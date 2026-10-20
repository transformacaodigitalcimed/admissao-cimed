// =====================================================================
// CIMED · Plataforma de Admissão · fase 1
// Site estático + Supabase (banco, login e regras de acesso).
// Nada para editar aqui: a configuração fica em config.js.
// =====================================================================
import { SUPABASE_URL, SUPABASE_ANON_KEY, LOGIN_MICROSOFT } from './config.js';

const api = window.API_FALSA || await criarApi();

// ================================================================ API
async function criarApi() {
  const { createClient } = await import('https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm');
  const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  const ok = ({ data, error }) => { if (error) throw new Error(traduzir(error.message)); return data; };
  const u = (s) => (s ? { email: s.user.email, id: s.user.id } : null);
  return {
    async sessao() { const { data } = await sb.auth.getSession(); return u(data.session); },
    // Nunca chamar o Supabase dentro do aviso de login: ele roda com uma trava
    // interna e a tela congela em "Carregando". Por isso o setTimeout.
    // A sessão inicial já é tratada em iniciar(); aqui só entrada e saída.
    aoMudar(cb) {
      sb.auth.onAuthStateChange((evt, s) => {
        if (evt === 'SIGNED_IN' || evt === 'SIGNED_OUT') setTimeout(() => cb(u(s)), 0);
      });
    },
    async entrar(email, senha) {
      const { error } = await sb.auth.signInWithPassword({ email, password: senha });
      if (error) throw new Error('E-mail ou senha não conferem.');
    },
    async entrarMicrosoft() {
      await sb.auth.signInWithOAuth({ provider: 'azure', options: { scopes: 'email', redirectTo: location.origin + location.pathname } });
    },
    async sair() { await sb.auth.signOut(); },
    async perfil(id) { return ok(await sb.from('perfis').select('*').eq('user_id', id).eq('ativo', true).maybeSingle()); },
    // a visão mascara CPF e tipo de PCD para quem não é Saúde, Admissão ou Remuneração
    async admissoes() { return ok(await sb.from('admissoes_visao').select('*').order('data_inicio')); },
    async umaAdmissao(id) { return ok(await sb.from('admissoes_visao').select('*').eq('id', id).single()); },
    async resumoDocs() { return ok(await sb.rpc('resumo_documentos')); },
    async criar(o) {
      const r = ok(await sb.from('admissoes').insert(o).select('id').single());
      return this.umaAdmissao(r.id);
    },
    async atualizar(id, p) {
      ok(await sb.from('admissoes').update(p).eq('id', id).select('id').single());
      return this.umaAdmissao(id);
    },
    async docs(id) { return ok(await sb.rpc('docs_da_admissao', { adm: id })); },
    async marcarDoc(id, tipo, status, obs) {
      return ok(await sb.from('documentos').update({ status, obs: obs || null }).eq('admissao_id', id).eq('tipo', tipo).select());
    },
    async eventos(id) { return ok(await sb.from('eventos').select('*').eq('admissao_id', id).order('em').order('id')); },
    async anotar(id, texto) { return ok(await sb.rpc('anotar', { adm: id, texto })); },
    async cobrar(id) { return ok(await sb.rpc('cobrar', { adm: id })); },
    async emails() { return ok(await sb.from('emails').select('*').order('criado_em', { ascending: false }).limit(300)); },
    async reenviar(id) { return ok(await sb.rpc('reenviar_email', { eid: id })); },
    // fila do WhatsApp (SQL 22): só existe depois do 22; quem chama trata a falha
    async whatsapp() { return ok(await sb.from('whatsapp_fila').select('*').order('criado_em', { ascending: false }).limit(200)); },
    async reenviarWhatsapp(id) { return ok(await sb.rpc('reenviar_whatsapp', { wid: id })); },
    // kits publicados on-line (Admissão e signatário leem): para a trilha "onde está cada kit"
    async kitsOnline() { return ok(await sb.from('kits_online').select('id,admissao_id,nome,versao,empresa_assinou,colab_assinou,token_hash,expira_em,cancelado_em,atualizado_em,blocos').is('cancelado_em', null).order('atualizado_em', { ascending: false }).limit(300)); },
    async areas() { return ok(await sb.from('areas_aviso').select('*').order('ordem')); },
    async salvarArea(id, p) { return ok(await sb.from('areas_aviso').update(p).eq('id', id).select()); },
    async locais() { return ok(await sb.from('locais').select('*').eq('ativo', true).order('nome')); },
    async parametros() { return ok(await sb.from('parametros').select('*')); },
    async entregas() { return ok(await sb.from('vw_entregas').select('*').order('data_inicio')); },
    async marcarEntrega(adm, area, por) { return ok(await sb.from('entregas').insert({ admissao_id: adm, area_id: area, por })); },
    async importar(linhas) { return ok(await sb.rpc('importar_admissoes', { linhas })); },

    // ---------- fase 2: arquivos e leitura por IA
    async arquivos(adm) {
      return ok(await sb.from('arquivos').select('id, tipo, caminho, nome, mime, status, lado, motivo, enviado_por, criado_em')
        .eq('admissao_id', adm).order('criado_em'));
    },
    // original = arquivo como veio (a impressão digital é dele); pronto = já reduzido para envio
    async subirArquivo(adm, tipo, original, pronto) {
      const sha = await sha256(original);
      const { data: igual } = await sb.from('arquivos').select('tipo').eq('admissao_id', adm).eq('sha256', sha).maybeSingle();
      if (igual) { const e = new Error('repetido'); e.repetidoEm = igual.tipo; throw e; }
      const ext = { 'application/pdf': 'pdf', 'image/png': 'png', 'image/webp': 'webp' }[pronto.type] || 'jpg';
      const caminho = `${adm}/${tipo.replace('#', '-')}/${crypto.randomUUID()}.${ext}`;
      const up = await sb.storage.from('documentos').upload(caminho, pronto, { contentType: pronto.type, upsert: false });
      if (up.error) throw new Error('Não consegui guardar o arquivo: ' + up.error.message);
      const { data, error } = await sb.from('arquivos')
        .insert({ admissao_id: adm, tipo, caminho, nome: original.name, mime: pronto.type, tamanho: pronto.size, sha256: sha })
        .select('id').single();
      if (error) { await sb.storage.from('documentos').remove([caminho]); throw new Error(traduzir(error.message)); }
      return data.id;
    },
    async lerArquivo(id) {
      const { data, error } = await sb.functions.invoke('ler-documento', { body: { arquivo_id: id } });
      if (error) {
        let msg = error.message;
        try { msg = (await error.context.json()).erro || msg; } catch { /* resposta sem corpo */ }
        throw new Error(msg);
      }
      return data;
    },
    async leituraFalhou(id, motivo) { return ok(await sb.rpc('leitura_falhou', { arq: id, motivo_falha: motivo })); },
    // abrir documento: registra quem abriu (sem o registro, o armazenamento não entrega)
    async verArquivo(id, caminho) {
      ok(await sb.rpc('registrar_acesso', { p_arquivo: id }));
      const { data, error } = await sb.storage.from('documentos').createSignedUrl(caminho, 60);
      if (error) throw new Error(traduzir(error.message));
      return data.signedUrl;
    },
    async removerArquivo(id, caminho) {
      ok(await sb.rpc('registrar_acesso', { p_arquivo: id }));
      ok(await sb.from('arquivos').delete().eq('id', id));
      // a linha já saiu; se o arquivo ficar no armazenamento, só registra (a ficha segue certa)
      const { error } = await sb.storage.from('documentos').remove([caminho]);
      if (error) console.warn('Arquivo ficou no armazenamento:', caminho, error.message);
      return true;
    },
    async acessos(adm) {
      return ok(await sb.from('acessos_arquivos').select('tipo, quem, em').eq('admissao_id', adm).order('em', { ascending: false }).limit(15));
    },
    // descarte automático (C2): o banco põe na fila, o site da Admissão apaga os arquivos
    async expurgoPendente() { return ok(await sb.rpc('expurgo_pendente')); },
    async apagarArquivos(bucket, caminhos) { const { error } = await sb.storage.from(bucket).remove(caminhos); if (error) throw new Error(error.message); return true; },
    async expurgoFeito() { return ok(await sb.rpc('expurgo_feito')); },

    // ---------- piloto da revisão de 19/10/2026 (SQL 13 a 16)
    async turmas() { return ok(await sb.from('turmas').select('*').order('data')); },
    async salvarTurma(o) { return ok(await sb.from('turmas').upsert(o).select('id')); },
    async apagarTurma(id) { return ok(await sb.from('turmas').delete().eq('id', id)); },
    async bps() { return ok(await sb.from('bps').select('*').order('unidade').order('area')); },
    // Segurança do Trabalho (SQL 20): matriz cargo → riscos/EPI/EPC e a Ordem de Serviço de cada pessoa
    async matrizOs() { return ok(await sb.from('matriz_os').select('*').eq('ativo', true).order('unidade').order('cargo')); },
    async salvarMatrizOs(o) { return ok(await sb.from('matriz_os').upsert(o).select('id')); },
    async apagarMatrizOs(id) { return ok(await sb.from('matriz_os').delete().eq('id', id)); },
    async osSalvar(adm, p, validar) { return ok(await sb.rpc('os_salvar', { adm, p, p_validar: !!validar })); },
    async osSugerir(adm) { return ok(await sb.rpc('os_sugerir', { adm })); },
    async asoDecidir(adm, decisao, p = {}) { return ok(await sb.rpc('aso_decidir', { adm, p_decisao: decisao, p_data: p.data || null, p_hora: p.hora || null, p_local: p.local || null })); },
    async salvarBp(o) { return ok(await sb.from('bps').upsert(o).select('id')); },
    async apagarBp(id) { return ok(await sb.from('bps').delete().eq('id', id)); },
    async destinatariosTurma() { return ok(await sb.from('destinatarios_turma').select('*').order('unidade')); },
    async salvarDestinatario(o) { return ok(await sb.from('destinatarios_turma').upsert(o).select('id')); },
    async apagarDestinatario(id) { return ok(await sb.from('destinatarios_turma').delete().eq('id', id)); },
    // a visão mascara o CPF do dependente para quem não é Saúde, Admissão ou Remuneração
    async dependentes(adm) { return ok(await sb.from('dependentes_visao').select('*').eq('admissao_id', adm).order('ordem')); },
    async salvarDependente(o) {
      const { id, idade: _i, ...campos } = o;
      if (id) return ok(await sb.from('dependentes').update(campos).eq('id', id).select('id'));
      return ok(await sb.from('dependentes').insert(campos).select('id'));
    },
    async apagarDependente(id) { return ok(await sb.from('dependentes').delete().eq('id', id)); },
    async listaTurma(data, locais, so) { return ok(await sb.rpc('lista_turma', { p_data: data, p_locais: locais, p_so_matriculados: so })); },
    async emailsListaTurma(locais) { return ok(await sb.rpc('emails_lista_turma', { p_locais: locais })); },
    async enviarListaTurma(data, locais, nome, base64) { return ok(await sb.rpc('enviar_lista_turma', { p_data: data, p_locais: locais, p_anexo_nome: nome || null, p_anexo_base64: base64 || null })); },
    async termos() { return ok(await sb.from('termos_transparencia').select('*').order('criado_em', { ascending: false })); },
    async termoTratar(id, parecer) { return ok(await sb.rpc('termo_tratar', { p_id: id, p_parecer: parecer })); },
    async termoRecebido(adm) { return ok(await sb.rpc('termo_recebido', { adm })); },
    async opcoesBeneficios(adm) { return ok(await sb.from('opcoes_beneficios').select('*').eq('admissao_id', adm).maybeSingle()); },
    async dossie(adm) { return ok(await sb.rpc('dossie_da_admissao', { adm })); },
    async dossieArquivado(adm, n, kit, resumo) { return ok(await sb.rpc('dossie_arquivado', { adm, p_arquivos: n, p_kit: kit, p_resumo: resumo })); },
    // baixa o arquivo em si (para o zip do dossiê): registra o acesso, como ao abrir
    async baixarArquivo(id, caminho) {
      ok(await sb.rpc('registrar_acesso', { p_arquivo: id }));
      const { data, error } = await sb.storage.from('documentos').download(caminho);
      if (error) throw new Error(traduzir(error.message));
      return data;
    },
    async baixarKitArquivo(caminho) {
      const { data, error } = await sb.storage.from('kits').download(caminho);
      if (error) throw new Error(traduzir(error.message));
      return data;
    },
  };
}

async function sha256(blob) {
  const h = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

// Foto vira JPG de até 2000 px (lê melhor, sobe mais rápido, custa menos na IA).
// PDF vai como está. HEIC e outros formatos são recusados com orientação.
async function prepararArquivo(f) {
  const LIMITE = 15 * 1024 * 1024;
  if (f.type === 'application/pdf' || /\.pdf$/i.test(f.name)) {
    if (f.size > LIMITE) throw new Error(`${f.name}: PDF maior que 15 MB.`);
    return f.type === 'application/pdf' ? f : new File([f], f.name, { type: 'application/pdf' });
  }
  if (/heic|heif/i.test(f.type + f.name)) throw new Error(`${f.name}: formato de foto do iPhone (HEIC). Peça a foto em JPG ou envie um print da tela.`);
  if (!/^image\/(jpeg|png|webp)$/.test(f.type)) throw new Error(`${f.name}: envie foto (JPG, PNG) ou PDF.`);
  const url = URL.createObjectURL(f); const img = new Image();
  try { await new Promise((ok, erro) => { img.onload = ok; img.onerror = erro; img.src = url; }); }
  catch { throw new Error(`${f.name}: não consegui abrir esta imagem.`); }
  finally { URL.revokeObjectURL(url); }
  const lado = Math.max(img.naturalWidth, img.naturalHeight);
  if (lado < 500) throw new Error(`${f.name}: imagem muito pequena (${img.naturalWidth}×${img.naturalHeight} px). O texto não vai dar para ler.`);
  const k = Math.min(1, 2000 / lado);
  const cv = document.createElement('canvas'); cv.width = Math.round(img.naturalWidth * k); cv.height = Math.round(img.naturalHeight * k);
  const x = cv.getContext('2d'); x.fillStyle = '#FFFFFF'; x.fillRect(0, 0, cv.width, cv.height); x.drawImage(img, 0, 0, cv.width, cv.height);
  const blob = await new Promise((r) => cv.toBlob(r, 'image/jpeg', 0.85));
  return new File([blob], f.name.replace(/\.\w+$/, '') + '.jpg', { type: 'image/jpeg' });
}

function traduzir(m) {
  if (/admissoes_cpf_ativo|duplicate key/i.test(m)) return 'Este CPF já está em uma admissão em andamento.';
  if (/admissoes_cpf_check/.test(m)) return 'CPF inválido: precisa ter 11 números.';
  if (/JSON object requested|0 rows|PGRST116/i.test(m)) return 'Você não tem permissão para esta alteração.';
  if (/row-level security/i.test(m)) return 'Você não tem permissão para esta ação.';
  if (/Failed to fetch|NetworkError/i.test(m)) return 'Sem conexão com o servidor. Verifique a internet e tente de novo.';
  return m;
}

// ========================================================= UTILIDADES
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const DIA = 864e5;
const dias = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
const iso = (d) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const deIso = (s) => { if (!s) return null; const [a, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(a, m - 1, d); };
const dm = (s) => { const d = typeof s === 'string' ? deIso(s) : s; return d ? String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') : '—'; };
const dmd = (s) => { const d = typeof s === 'string' ? deIso(s) : s; return d ? dias[d.getDay()] + ' ' + dm(d) : '—'; };
const quando = (t) => { const d = new Date(t); return dm(d) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };
const corteDe = (inicio) => { const c = deIso(inicio); c.setDate(c.getDate() - 4); c.setHours(12, 0, 0, 0); return c; };
// corte de uma pessoa: o da turma cadastrada (vem na visão), senão 4 dias antes ao meio-dia
const corteA = (a) => (a && a.corte_documentos ? new Date(a.corte_documentos) : corteDe(a.data_inicio));
const hm = (d) => String(d.getHours()).padStart(2, '0') + 'h' + (d.getMinutes() ? String(d.getMinutes()).padStart(2, '0') : '');
const dmh = (t) => { const d = t instanceof Date ? t : new Date(t); return dmd(d) + ' ' + hm(d); };
// "2026-10-20T12:00" local → ISO com fuso; e o contrário, para o input datetime-local
const deLocal = (s) => (s ? new Date(s).toISOString() : null);
const paraLocal = (t) => { if (!t) return ''; const d = new Date(t); return iso(d) + 'T' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };
function baixarBlob(nome, blob) { const u = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = u; a.download = nome; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 4000); }
// .zip sem compressão (só guarda), para o dossiê: nome em UTF-8, CRC-32, data de agora
const TCRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = TCRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zipar(entradas) {
  const enc = new TextEncoder(); const partes = []; const centro = []; let off = 0;
  const d = new Date(); const hora = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1); const data = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  const u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; const u16 = (v) => [v & 255, (v >>> 8) & 255];
  entradas.forEach(({ nome, bytes }) => {
    const n = enc.encode(nome); const crc = crc32(bytes);
    const loc = new Uint8Array([...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(hora), ...u16(data), ...u32(crc), ...u32(bytes.length), ...u32(bytes.length), ...u16(n.length), ...u16(0), ...n]);
    partes.push(loc, bytes);
    centro.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(hora), ...u16(data), ...u32(crc), ...u32(bytes.length), ...u32(bytes.length), ...u16(n.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off), ...n]));
    off += loc.length + bytes.length;
  });
  const tam = centro.reduce((s, c) => s + c.length, 0);
  const fim = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(entradas.length), ...u16(entradas.length), ...u32(tam), ...u32(off), ...u16(0)]);
  return new Blob([...partes, ...centro, fim], { type: 'application/zip' });
}
const nomeArquivoSeguro = (s) => semAcento(s).replace(/[^A-Z0-9 ._-]/g, ' ').replace(/\s+/g, ' ').trim();
const plural = (n, s, p) => n + ' ' + (n === 1 ? s : p);
const titulo = (n) => String(n || '').toLowerCase().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()).replace(/\b(Da|De|Do|Das|Dos|E)\b/g, (w) => w.toLowerCase());
const primeiro = (n) => titulo(String(n || '').split(/\s+/)[0]);
const semAcento = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
function cpfValido(c) {
  c = String(c).replace(/\D/g, ''); if (c.length !== 11 || /^(\d)\1+$/.test(c)) return false;
  for (const t of [9, 10]) { let s = 0; for (let i = 0; i < t; i++) s += +c[i] * (t + 1 - i); let r = s * 10 % 11; if (r === 10) r = 0; if (r !== +c[t]) return false; }
  return true;
}
const fmtCpf = (c) => { c = String(c || '').padStart(11, '0'); return c.slice(0, 3) + '.' + c.slice(3, 6) + '.' + c.slice(6, 9) + '-' + c.slice(9); };
function idade(n) { if (!n) return null; const d = deIso(n), h = new Date(); let a = h.getFullYear() - d.getFullYear(); const m = h.getMonth() - d.getMonth(); if (m < 0 || (m === 0 && h.getDate() < d.getDate())) a--; return a; }

const ICO = {
  ok: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8.5l3 3 7-7"/></svg>',
  at: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M8 2l6.5 11.5h-13z"/><path d="M8 6.5v3M8 11.4v.1" stroke-linecap="round"/></svg>',
  cr: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M4 4l8 8M12 4l-8 8"/></svg>',
  ag: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="8" cy="8" r="6"/><path d="M8 4.5V8l2.5 1.5"/></svg>',
  info: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="8" cy="8" r="6"/><path d="M8 7.5v3.5M8 5v.1"/></svg>',
};
const chip = (cls, ico, txt) => `<span class="chip ${cls}">${ICO[ico] || ''}${esc(txt)}</span>`;
const svgI = (p) => `<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
const MI = {
  inicio: svgI('<path d="M3 11l9-7 9 7v9a1 1 0 01-1 1h-5v-6H9v6H4a1 1 0 01-1-1z"/>'),
  novo: svgI('<circle cx="10" cy="8" r="4"/><path d="M3 20c0-3.5 3-6 7-6s7 2.5 7 6M19 8v6M16 11h6"/>'),
  painel: svgI('<rect x="3" y="4" width="5" height="16" rx="1.5"/><rect x="10" y="4" width="5" height="11" rx="1.5"/><rect x="17" y="4" width="4" height="7" rx="1.5"/>'),
  docs: svgI('<path d="M7 3h7l5 5v13H7z"/><path d="M14 3v5h5M10 13l2 2 4-4"/>'),
  esp: svgI('<circle cx="7" cy="8" r="3"/><circle cx="17" cy="8" r="3"/><path d="M2 20c0-3 2.2-5 5-5s5 2 5 5M12 20c0-3 2.2-5 5-5s5 2 5 5"/>'),
  mail: svgI('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>'),
  saude: svgI('<path d="M12 21s-8-4.5-8-11a4.5 4.5 0 018-2.8A4.5 4.5 0 0120 10c0 6.5-8 11-8 11z"/><path d="M12 9v5M9.5 11.5h5"/>'),
  remun: svgI('<rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/>'),
  kit: svgI('<path d="M4 19l4-1 10-10-3-3L5 15z"/><path d="M14 6l3 3M4 22h16"/>'),
};

const ETAPAS = [
  { id: 'aguardando', n: 'Aguardando candidato', d: 'Nenhum documento recebido ainda' },
  { id: 'pendencia', n: 'Pendência com candidato', d: 'Faltou documento ou precisa reenviar' },
  { id: 'conferencia', n: 'Para conferir', d: 'Tudo recebido, a analista confere' },
  { id: 'docs_ok', n: 'Documentação completa', d: 'Áreas avisadas; falta ASO, remuneração ou LG' },
  { id: 'registro', n: 'Pronto para registro', d: 'Docs, ASO apto, remuneração ok e LG movimentado' },
  { id: 'assinatura', n: 'Matriculado · assinatura', d: 'Kit gerado, falta assinar' },
  { id: 'pronto', n: 'Pronto para o Dia 1', d: 'Kit assinado' },
];
const ETAPA = Object.fromEntries(ETAPAS.map((e) => [e.id, e]));
const GATILHO = { cadastro: 'Quando o R&S cadastra', docs_ok: 'Quando a documentação fica completa', matricula: 'Quando a matrícula é gerada', declinou: 'Quando a admissão é cancelada (declínio ou não comparecimento)', termo: 'Quando o Termo de Transparência tem resposta "Sim"' };
const TIPO_EMAIL = { cadastro: 'Novo aprovado', boasvindas: 'Boas-vindas', cobranca: 'Cobrança', docs_ok: 'Documentação completa', matricula: 'Matrícula', lg: 'Movimentar no LG', convite: 'Convite da integração', turma: 'Lista da turma', beneficios_opcoes: 'Opções de benefícios', termo: 'Termo de Transparência', via_assinada: 'Via assinada do kit', aso: 'Exame admissional (ASO)', declinou: 'Admissão cancelada' };
// a jornada de e-mails do colaborador, na ordem (os demais e-mails vão para as áreas e a equipe)
const JORNADA_EMAILS = [
  ['boasvindas', 'Cadastro do aprovado pelo R&S', 'Boas-vindas: data da integração, lista de documentos, como enviar (LG), prazo do corte da turma, link do Código de Conduta e aviso de privacidade.'],
  ['cobranca', 'Todo dia, enquanto faltar documento (até o corte da turma)', 'Cobrança dos documentos que faltam ou precisam ser reenviados; também pelo botão "Cobrar" da ficha.'],
  ['convite', 'Matrícula registrada', 'Convite da integração: data, hora, local (ou link do Teams) e quem procurar ao chegar; gestor e recrutadora em cópia.'],
  ['via_assinada', 'Colaborador assina o kit e escolhe e-mail', 'Via assinada do kit admissional em PDF, com o código de verificação.'],
];
// e-mails que dependem de um acontecimento
const EMAILS_EVENTO = [
  ['aso', 'Saúde agenda o exame (ASO agendado com data)', 'Ao candidato: data, hora, local e orientações do exame (parâmetro aso_orientacoes). Se o ASO fica inapto, a analista de R&S e a Admissão recebem o aviso.'],
  ['declinou', 'Admissão cancelada', 'Às áreas com aviso no gatilho "admissão cancelada" (TI: cancelar chamado, acessos e equipamento).'],
];
const EMAILS_EQUIPE = [
  ['cadastro', 'Cadastro', 'Áreas com aviso no cadastro (Saúde para agendar o ASO, Admissão, Remuneração, Sindical no Comercial, TI previsão de equipamento).'],
  ['lg', 'ASO apto', 'Pedido à analista de R&S para movimentar a pessoa no LG.'],
  ['docs_ok', 'Documentação completa', 'Saúde (confirmar ASO), Frotas, Facilities, R&S, BP da unidade e gestor.'],
  ['matricula', 'Matrícula', 'Áreas da matrícula (TI para abrir o chamado, Benefícios com dependentes, Segurança, DP).'],
  ['turma', 'Botão "Enviar por e-mail" da tela Turma e lista', 'Lista da turma (U1/U2) para os destinatários da unidade, requisitantes em cópia.'],
  ['beneficios_opcoes', 'Kit assinado', 'Escolhas de benefícios feitas no kit, para a área de Benefícios.'],
  ['termo', 'Kit assinado com algum "Sim" no termo', 'Aviso ao Compliance sobre o Termo de Transparência.'],
];
const PARENTESCOS = ['Filho(a)', 'Enteado(a)', 'Cônjuge', 'Companheiro(a)', 'Pai', 'Mãe'];
const ESOCIAL = ['S-2200 (admissão)', 'S-2300 (estágio, sem vínculo)', 'S-2200 retificação', 'S-3000 (exclusão)'];
const GRUPO_TURMA = { todos: 'Todas as unidades', pa_sp_comercial: 'PA / SP / Comercial', ssbv_cds: 'SSBV e CDs' };
const TERMO_PERGUNTAS = [
  ['q1', 'Possui familiar de 1º grau (pais, filhos, irmãos, cônjuge ou companheiro(a)) trabalhando na Cimed?', 'Nome, parentesco, cargo, área e turno'],
  ['q2', 'Mantém relacionamento amoroso com alguém da mesma área ou sob a mesma gestão?', 'Nome e área'],
  ['q3', 'Presta serviço a órgão público?', 'Tipo de serviço, órgão, cargo e carga horária'],
  ['q4', 'Possui familiar de 1º grau que seja agente público ou político?', 'Nome, parentesco, órgão e cargo'],
  ['q5', 'Possui familiar ou relacionamento em empresa concorrente da Cimed?', 'Nome, vínculo e empresa'],
  ['q6', 'Possui familiar ou relacionamento em empresa que tem negócios ou está em negociação com a Cimed?', 'Nome, vínculo e empresa'],
  ['q7', 'Possui familiar ou relacionamento com participação em empresa em litígio contra a Cimed?', 'Nome, vínculo e empresa'],
  ['q8', 'Tem participação societária em empresa que tem negócios com a Cimed?', 'Nome da empresa, CNPJ e ramo'],
  ['q9', 'Tem participação societária em qualquer outra empresa?', 'Nome da empresa, CNPJ e ramo'],
  ['q10', 'Presta serviço a outra empresa?', 'Nome da empresa, CNPJ, função e carga horária'],
  ['q11', 'Há outra situação que possa gerar conflito de interesses?', 'Descreva'],
];
const CONDICAO = { todos: 'Todos', frota: 'Só quem usa veículo da frota', ti: 'Só quem precisa de equipamento', comercial: 'Só comercial', fabrica_cd: 'Fábricas e CDs', fabrica: 'Fábricas' };
const ST_DOC = { falta: 'Não enviado', analise: 'Recebido, conferir', reenviar: 'Pedir novo envio', ok: 'Conferido' };

// ============================================================= ESTADO
const E = {
  sessao: null, perfil: null, view: null, carregado: false,
  admissoes: [], resumo: {}, locais: [], areas: [], params: {}, emails: [], whatsapp: [], kitsOnline: [], entregas: [],
  turma: null, localF: '', modo: 'quadro', selDoc: null, selMail: null, previa: null,
  turmasCal: [], bps: [], dest: [], termos: [], depsNovo: [], turmaLista: null, locaisLista: null, soMatriculados: true, selTermo: null, matrizOs: [],
  verComo: null,
};
const param = (k, padrao = '') => (E.params[k] == null ? padrao : String(E.params[k]));
const ligado = (k, padrao = 'sim') => param(k, padrao) === 'sim';
const localDe = (nome) => E.locais.find((l) => l.nome === nome) || {};
// "Ver como" (admin, 20/10/2026): o admin escolhe um perfil e a plataforma mostra só o que esse perfil vê.
// Muda só a tela e os botões; os dados continuam vindo com a permissão do admin.
const PERFIS_VER_COMO = [
  ['admin', 'Admin (tudo)', ['admin'], null],
  ['rs', 'R&S (recrutadora)', ['rs'], null],
  ['admissao', 'Admissão · ADM de Pessoal', ['admissao'], null],
  ['saude', 'Saúde ocupacional', ['saude'], null],
  ['remuneracao', 'Remuneração', ['remuneracao'], null],
  ['signatario', 'Signatário (Keila)', ['signatario'], null],
  ['compliance', 'Compliance', ['compliance'], null],
  ['esp-seguranca', 'Especialista · Segurança do Trabalho', ['especialista'], 'seguranca'],
  ['esp-ti', 'Especialista · TI', ['especialista'], 'ti'],
  ['esp-facilities', 'Especialista · Facilities', ['especialista'], 'facilities'],
];
const souAdminDeVerdade = () => !!E.perfil && (E.perfil.papeis || []).includes('admin');
const papeisAtivos = () => (E.verComo ? E.verComo.papeis : (E.perfil ? E.perfil.papeis : []));
const tem = (...p) => !!E.perfil && (papeisAtivos().includes('admin') || p.some((x) => papeisAtivos().includes(x)));
function definirVerComo(chave) {
  const def = PERFIS_VER_COMO.find((x) => x[0] === chave);
  E.verComo = def && def[0] !== 'admin' && souAdminDeVerdade() ? { chave: def[0], rotulo: def[1], papeis: def[2], area_id: def[3] } : null;
  try { sessionStorage.setItem('cimed.verComo', E.verComo ? E.verComo.chave : ''); } catch (e) { /* sem sessionStorage */ }
}
// a primeira tela de cada perfil: quem não é R&S nem Admissão abre direto na própria
function visaoInicial() {
  const so = !tem('rs', 'admissao');
  return so && tem('saude') ? 'saude' : so && tem('remuneracao') ? 'remuneracao' : so && tem('signatario') ? 'kit' : so && tem('compliance') ? 'compliance' : navVisivel()[0]?.[0];
}

// [id, nome, ícone, quem vê, seção do menu]
const NAV = [
  ['inicio', 'Início', 'inicio', () => tem('rs', 'admissao'), 1],
  ['novo', 'Cadastrar aprovado', 'novo', () => tem('rs', 'admissao'), 1],
  ['painel', 'Painel da admissão', 'painel', () => tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario'), 1],
  ['turma', 'Turma e lista', 'esp', () => tem('rs', 'admissao'), 1],
  ['documentos', 'Documentos', 'docs', () => tem('admissao'), 1],
  ['compliance', 'Termo de Transparência', 'docs', () => tem('compliance'), 1],
  ['kit', 'Esteira de assinatura', 'kit', () => tem('admissao', 'signatario'), 1],
  ['remuneracao', 'Remuneração', 'remun', () => tem('remuneracao', 'admissao'), 1],
  ['saude', 'Saúde ocupacional', 'saude', () => tem('saude'), 1],
  ['seguranca', 'Segurança do Trabalho', 'esp', () => souSeguranca(), 1],
  ['ti', 'TI · equipamentos e acessos', 'esp', () => souTi(), 1],
  ['entregas', 'Minhas entregas', 'esp', () => tem('especialista', 'admissao'), 1],
  ['especialistas', 'Especialistas e avisos', 'esp', () => tem('admissao'), 2],
  ['emails', 'Fila de e-mails', 'mail', () => tem('admissao'), 2],
];
const navVisivel = () => NAV.filter((n) => n[3]());

// Quem recebe aviso num momento, pelas mesmas regras do banco (só para o texto da tela)
function avisadas(gatilho, a) {
  if (!a) return '';
  const tipo = E.locais.find((l) => l.nome === a.local)?.tipo || '';
  const vale = (ar) => ar.ativo && ar.gatilho === gatilho && (ar.emails || []).length &&
    !(ar.condicao === 'frota' && !a.usa_frota) && !(ar.condicao === 'ti' && !(a.ti_itens || []).length) &&
    !(ar.condicao === 'comercial' && tipo !== 'comercial') && !(ar.condicao === 'fabrica_cd' && !['fabrica', 'cd'].includes(tipo)) &&
    !(ar.condicao === 'fabrica' && tipo !== 'fabrica');
  const n = E.areas.filter(vale).map((ar) => ar.nome.split(' ·')[0]);
  return n.length ? 'Avisos na fila para: ' + n.join(', ') + '.' : '';
}

// ======================================================== CICLO
let toastT;
function toast(msg, erro) {
  clearTimeout(toastT);
  let t = $('#toast');
  if (!t) { t = document.createElement('div'); t.id = 'toast'; t.className = 'toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
  t.innerHTML = (erro ? ICO.at : ICO.ok).replace('<svg', '<svg width="16" height="16"') + '<span>' + esc(msg) + '</span>';
  t.hidden = false; toastT = setTimeout(() => (t.hidden = true), erro ? 6000 : 3800);
}
async function tentar(fn, sucesso) {
  try { const r = await fn(); if (sucesso) toast(sucesso); return r; }
  catch (e) { toast(e.message || String(e), true); return undefined; }
}

async function iniciar() {
  api.aoMudar(async (s) => {
    const antes = E.sessao?.id; E.sessao = s;
    if ((s?.id || null) !== (antes || null)) await aposLogin();
  });
  E.sessao = await api.sessao();
  await aposLogin();
}

async function aposLogin() {
  E.perfil = null; E.carregado = false;
  if (!E.sessao) return telaLogin();
  E.perfil = await tentar(() => api.perfil(E.sessao.id));
  if (!E.perfil) return telaSemAcesso();
  E.perfil.papeis = E.perfil.papeis || [];
  try { definirVerComo(sessionStorage.getItem('cimed.verComo') || ''); } catch (e) { E.verComo = null; }
  E.view = visaoInicial();
  if (!E.view) return telaSemAcesso();
  await carregar();
  casca();
  if (tem('admissao')) limparExpurgo();
}

// Arquivos que o descarte automático (30 dias) mandou apagar
async function limparExpurgo() {
  try {
    const l = (api.expurgoPendente && await api.expurgoPendente()) || [];
    if (!l.length) return;
    const porBucket = {};
    l.forEach((x) => { (porBucket[x.bucket] = porBucket[x.bucket] || []).push(x.caminho); });
    for (const [b, cs] of Object.entries(porBucket)) await api.apagarArquivos(b, cs);
    await api.expurgoFeito();
  } catch (e) { console.warn('Descarte automático não terminou:', e); }
}

async function carregar() {
  const tarefas = [];
  if (tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario') || souSeguranca() || souTi()) {
    tarefas.push(api.admissoes().then((r) => (E.admissoes = r || [])));
    if (tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario')) tarefas.push(api.resumoDocs().then((r) => (E.resumo = Object.fromEntries((r || []).map((x) => [x.admissao_id, x])))));
  }
  if (souSeguranca() && api.matrizOs) tarefas.push(api.matrizOs().then((r) => (E.matrizOs = r || [])).catch(() => (E.matrizOs = [])));
  tarefas.push(api.locais().then((r) => (E.locais = r || [])));
  tarefas.push(api.areas().then((r) => (E.areas = r || [])));
  tarefas.push(api.parametros().then((r) => (E.params = Object.fromEntries((r || []).map((x) => [x.chave, x.valor])))));
  if (tem('admissao')) tarefas.push(api.emails().then((r) => (E.emails = r || [])));
  if (tem('admissao') && api.whatsapp) tarefas.push(api.whatsapp().then((r) => (E.whatsapp = r || [])).catch(() => (E.whatsapp = [])));
  if (tem('admissao', 'signatario') && api.kitsOnline) tarefas.push(api.kitsOnline().then((r) => (E.kitsOnline = r || [])).catch(() => (E.kitsOnline = [])));
  if (tem('especialista', 'admissao')) tarefas.push(api.entregas().then((r) => (E.entregas = r || [])));
  if (tem('rs', 'admissao')) {
    tarefas.push(api.turmas().then((r) => (E.turmasCal = r || [])));
    tarefas.push(api.bps().then((r) => (E.bps = r || [])));
    tarefas.push(api.destinatariosTurma().then((r) => (E.dest = r || [])));
  }
  if (tem('compliance')) tarefas.push(api.termos().then((r) => (E.termos = r || [])));
  await tentar(() => Promise.all(tarefas));
  if (!E.turma) {
    const futuras = turmas().filter((t) => deIso(t) >= deIso(iso(new Date())));
    E.turma = futuras[0] || 'todas';
  }
  E.carregado = true;
}
async function recarregar() { await carregar(); render(); }

const vivas = () => E.admissoes.filter((a) => !a.declinou_em);
const turmas = () => [...new Set(E.admissoes.filter((a) => !a.declinou_em).map((a) => a.data_inicio))].sort();
const resumo = (a) => E.resumo[a.id] || { total: 0, ok: 0, analise: 0, reenviar: 0, pendentes: [] };
const adm = (id) => E.admissoes.find((a) => a.id === id);
const foraCorte = (a) => ['aguardando', 'pendencia', 'conferencia'].includes(a.etapa) && Date.now() > corteA(a).getTime();
const risco = (a) => ['aguardando', 'pendencia', 'conferencia'].includes(a.etapa) && Date.now() > corteA(a).getTime() - DIA;
// corte de documentos de uma turma (data): o da turma do calendário, senão o de alguém dela, senão a regra fixa
function corteTurma(turma) {
  const t = E.turmasCal.find((x) => x.data === turma && x.ativo);
  if (t) return new Date(t.corte_documentos);
  const a = vivas().find((x) => x.data_inicio === turma && x.corte_documentos);
  return a ? new Date(a.corte_documentos) : corteDe(turma);
}
const faltaLg = (a) => ligado('exige_lg_movimentado') && !a.declinou_em && !a.matricula && (a.aso_status === 'apto' || a.aso_status === 'apto_restricao') && !a.lg_movimentado_em;
const faltaDossie = (a) => !a.declinou_em && a.kit_status === 'concluido' && !a.arquivado_em && !a.expurgado_em;
// kit em blocos por área (decisão 4 de 19/10/2026): status gravado pela Esteira em admissoes.kit_blocos
// (cópia de kits_online.blocos). Cada bloco tem o seu signatário pela Cimed; o colaborador assina tudo.
const AREAS_KIT = { dp: 'DP', beneficios: 'Benefícios', compliance: 'Compliance', seguranca: 'Segurança' };
const blocosKit = (a) => (a.kit_blocos && typeof a.kit_blocos === 'object') ? Object.keys(AREAS_KIT).filter((k) => a.kit_blocos[k]).map((k) => Object.assign({ area: k, titulo: AREAS_KIT[k] }, a.kit_blocos[k])) : [];
const faltaAssinar = (a) => blocosKit(a).filter((b) => b.assinavel !== false && !b.assinado_empresa_em).map((b) => b.titulo);
// área sem ninguém cadastrado em signatarios_por_area: o bloco fica pendente até a Josianny informar o nome
function semSignatario(area) {
  try { const m = JSON.parse(E.params.signatarios_por_area || '{}'); return !!m && typeof m === 'object' && Object.keys(m).length > 0 && !(Array.isArray(m[area]) && m[area].length); }
  catch (e) { return false; }
}
function faixaBlocos() {
  const pend = vivas().filter((a) => a.matricula && a.kit_status !== 'concluido' && faltaAssinar(a).length);
  if (!pend.length) return '';
  const por = {}; pend.forEach((a) => faltaAssinar(a).forEach((t) => { (por[t] = por[t] || []).push(titulo(a.nome)); }));
  return `<div class="faixa"><p><b>Falta a Cimed assinar:</b> ${Object.entries(por).map(([t, ns]) => `${esc(t)} em ${plural(ns.length, 'kit', 'kits')} (${ns.map(esc).join(', ')})`).join(' · ')}. Hoje só o DP leva a assinatura da Cimed; em Benefícios e Compliance só o colaborador assina.</p></div>`;
}

// ============================================================ TELAS FIXAS
// ícones do olhinho da senha (aberto = senha escondida, clique para mostrar; fechado = senha à vista)
const OLHO = {
  aberto: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>',
  fechado: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17.9 17.9A10.5 10.5 0 0 1 12 19c-6.5 0-10-7-10-7a18 18 0 0 1 5.1-5.9"/><path d="M9.9 4.2A10.6 10.6 0 0 1 12 4c6.5 0 10 7 10 7a18 18 0 0 1-2.2 3.2"/><path d="M14.1 14.1a3 3 0 0 1-4.2-4.2"/><line x1="2" y1="2" x2="22" y2="22"/></svg>',
};
function telaLogin(msg) {
  $('#raiz').innerHTML = `<div class="entrada"><form class="caixa" id="fLogin">
    <span class="marca-g">CIMED</span><h1>Plataforma de Admissão</h1>
    <p class="apoio">Entre com o seu e-mail da Cimed. O acesso é liberado pela Transformação Digital do RH.</p>
    ${LOGIN_MICROSOFT ? `<button type="button" class="bt pri grande" data-act="microsoft">Entrar com a conta Microsoft da Cimed</button><div class="sep"></div>` : ''}
    <div class="campo"><label for="lEmail">E-mail</label><input id="lEmail" type="email" autocomplete="username" required></div>
    <div class="campo"><label for="lSenha">Senha</label><div class="senha-olho"><input id="lSenha" type="password" autocomplete="current-password" required><button type="button" id="lOlho" aria-label="Mostrar senha" aria-pressed="false" title="Mostrar senha">${OLHO.aberto}</button></div></div>
    ${msg ? `<span class="erro" style="color:var(--cr);font-weight:600">${esc(msg)}</span>` : ''}
    <button class="bt ${LOGIN_MICROSOFT ? '' : 'pri'} grande" type="submit">Entrar</button>
  </form></div>`;
  // olhinho: mostra ou esconde a senha digitada
  $('#lOlho').addEventListener('click', () => {
    const s = $('#lSenha'), b = $('#lOlho'), ver = s.type === 'password';
    s.type = ver ? 'text' : 'password';
    b.innerHTML = ver ? OLHO.fechado : OLHO.aberto;
    b.setAttribute('aria-pressed', String(ver));
    b.title = b.ariaLabel = ver ? 'Esconder senha' : 'Mostrar senha';
    s.focus();
  });
  $('#fLogin').addEventListener('submit', async (e) => {
    e.preventDefault();
    try { await api.entrar($('#lEmail').value.trim(), $('#lSenha').value); E.sessao = await api.sessao(); await aposLogin(); }
    catch (err) { telaLogin(err.message); }
  });
}
function telaSemAcesso() {
  $('#raiz').innerHTML = `<div class="entrada"><div class="caixa"><span class="marca-g">CIMED</span><h1>Acesso não liberado</h1>
    <p class="apoio">A conta ${esc(E.sessao?.email || '')} entrou, mas não está na lista de pessoas da Plataforma de Admissão. Peça a liberação à Transformação Digital do RH.</p>
    <button class="bt pri" data-act="sair">Sair</button></div></div>`;
}
function casca() {
  const podeBuscar = tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario');
  $('#raiz').innerHTML = `<div class="casca">
    <nav class="menu" id="menu" aria-label="Módulos"></nav>
    <div class="conteudo">
      <header class="topo">
        ${podeBuscar ? `<div class="busca"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
          <input id="busca" type="search" placeholder="Buscar pessoa pelo nome ou CPF" autocomplete="off" aria-label="Buscar pessoa"><div class="busca-res" id="buscaRes" hidden></div></div>` : '<div class="esp"></div>'}
        ${E.params.envio_ativo !== 'sim' && tem('admissao') ? '<span class="selo-homolog" title="Os e-mails ficam na fila até o envio ser ligado">Homologação</span>' : ''}
        ${tem('rs', 'admissao') ? `<button class="bt" data-act="ir" data-v="novo">${MI.novo} Cadastrar aprovado</button>` : ''}
        <div class="topo-dir">
        ${souAdminDeVerdade() ? `<label class="ver-como" title="Veja a plataforma como cada perfil de acesso vê"><span>Ver como</span><select data-change="verComo">${PERFIS_VER_COMO.map(([c, r]) => `<option value="${c}" ${(E.verComo ? E.verComo.chave : 'admin') === c ? 'selected' : ''}>${esc(r)}</option>`).join('')}</select></label>` : ''}
        ${sinoHtml()}
        <div class="quem"><span class="avatar">${esc((E.perfil.nome || '?').slice(0, 2).toUpperCase())}</span><span class="quem-nome">${esc(E.perfil.nome)}</span><button class="bt peq" data-act="sair">Sair</button></div>
        </div>
      </header>
      ${E.verComo ? `<div class="aviso-faixa info ver-como-faixa">${ICO.info.replace('<svg', '<svg width="16" height="16"')}<span style="flex:1">Você está vendo a plataforma como <b>${esc(E.verComo.rotulo)}</b>. Só as telas e os botões mudam; os dados continuam os seus.</span><button class="bt peq" data-act="verComoSair">Voltar ao admin</button></div>` : ''}
      <main id="main"></main>
      <div id="esteiraHost" class="esteira-host" hidden></div>
    </div></div>`;
  $('#busca')?.addEventListener('input', buscar);
  $('#busca')?.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.target.value = ''; buscar(); } });
  render();
  iniciarSino();
}

// ============================================================ SINO (20/10/2026)
// Pendências da sua parte e novidades que chegaram, por perfil. O número é o que você ainda não abriu.
const SVG_SINO = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.7 21a2 2 0 0 1-3.4 0"/></svg>';
const chaveVistos = () => 'cimed.sino.vistos.' + String((E.perfil && E.perfil.email) || '').toLowerCase();
function vistosSino() { try { return JSON.parse(localStorage.getItem(chaveVistos()) || '{}') || {}; } catch (e) { return {}; } }
function marcarVistos(chaves) {
  const v = vistosSino(); const agora = Date.now();
  chaves.forEach((c) => { v[c] = agora; });
  // esquece o que já não está na lista há 60 dias, para não crescer sem parar
  Object.keys(v).forEach((c) => { if (agora - v[c] > 60 * 864e5) delete v[c]; });
  try { localStorage.setItem(chaveVistos(), JSON.stringify(v)); } catch (e) { /* sem espaço */ }
}
// áreas em que quem está logado assina pela Cimed (e-mail em signatarios_por_area, SQL 23)
function minhasAreasSig() {
  try { const m = JSON.parse(E.params.signatarios_por_area || '{}'); const eu = String((E.perfil && E.perfil.email) || '').toLowerCase(); return Object.keys(m).filter((a) => (m[a] || []).some((s) => String(s.email || '').toLowerCase() === eu)); }
  catch (e) { return []; }
}
function notificacoes() {
  const l = []; const nome = (n) => titulo(n || '');
  const recente = (k) => !k.atualizado_em || (Date.now() - new Date(k.atualizado_em).getTime()) < 14 * 864e5;
  const areas = tem('signatario') ? minhasAreasSig() : [];
  (E.kitsOnline || []).forEach((k) => {
    if (k.cancelado_em) return;
    const b = k.blocos || {};
    if (areas.length && areas.some((a) => b[a] && b[a].assinavel !== false && !b[a].assinado_empresa_em)) l.push({ chave: 'sig:' + k.id + ':' + k.versao, nivel: 0, texto: 'Chegou para você assinar: ' + nome(k.nome), view: 'kit' });
    if (!tem('admissao') || !recente(k)) return;
    if (k.colab_assinou) l.push({ chave: 'colab:' + k.id, nivel: 1, texto: nome(k.nome) + ' assinou o kit', view: 'kit', id: k.admissao_id });
    else if (k.empresa_assinou) l.push({ chave: 'cimed:' + k.id + ':' + k.versao, nivel: 2, texto: 'A Cimed assinou o kit de ' + nome(k.nome) + ': pendente o colaborador assinar', view: 'kit', id: k.admissao_id });
  });
  if (tem('admissao')) (E.emails || []).filter((m) => m.status === 'erro').forEach((m) => l.push({ chave: 'mail:' + m.id, nivel: 1, texto: 'E-mail com erro: ' + (m.assunto || ''), view: 'emails' }));
  if (tem('especialista') && !tem('admissao')) (E.entregas || []).filter((x) => !x.pronto_em).slice(0, 20).forEach((x) => l.push({ chave: 'entrega:' + x.admissao_id + ':' + x.area_id, nivel: 2, texto: 'Entrega pendente da sua área: ' + nome(x.nome) + (x.area_nome ? ' (' + x.area_nome + ')' : ''), view: 'entregas' }));
  // a lista de atenção do Início é de quem toca a admissão; quem só assina ou só vê o Compliance não recebe
  if (tem('rs', 'admissao', 'saude', 'remuneracao') || souSeguranca() || souTi()) listaAtencao().forEach(([p, a, tx]) => l.push({ chave: 'at:' + a.id + ':' + tx.slice(0, 24), nivel: p, texto: nome(a.nome) + ' · ' + tx, id: a.id }));
  l.sort((x, y) => x.nivel - y.nivel);
  return l;
}
function sinoHtml() {
  const l = notificacoes(); const v = vistosSino(); const novos = l.filter((x) => !v[x.chave]).length;
  return `<div class="sino-wrap"><button class="sino${novos ? ' aceso' : ''}" data-act="sino" aria-label="Pendências e novidades" title="${novos ? plural(novos, 'novidade', 'novidades') : 'Pendências e novidades'}">${SVG_SINO}${novos ? `<span class="sino-n num">${novos > 99 ? '99+' : novos}</span>` : ''}</button><div class="sino-lista" id="sinoLista" hidden></div></div>`;
}
function atualizarSino() { const w = $('.sino-wrap'); if (w) w.outerHTML = sinoHtml(); }
function abrirSino() {
  const box = $('#sinoLista'); if (!box) return;
  if (!box.hidden) { box.hidden = true; return; }
  const l = notificacoes(); const v = vistosSino();
  box.innerHTML = l.length
    ? l.map((x) => `<button data-act="sinoIr" data-v="${x.view || ''}" data-id="${x.id || ''}">${x.nivel === 0 ? chip('c-cr', 'cr', '!') : ''}<span style="flex:1">${esc(x.texto)}</span>${v[x.chave] ? '' : '<span class="novo">novo</span>'}</button>`).join('')
    : '<div class="vazio-sino">Nada pendente da sua parte. Quando chegar algo, o sino acende.</div>';
  box.hidden = false;
  marcarVistos(l.map((x) => x.chave));
  const b = $('.sino'); if (b) { b.classList.remove('aceso'); const n = b.querySelector('.sino-n'); if (n) n.remove(); }
}
let SINO_TIMER = null;
function iniciarSino() {
  if (SINO_TIMER) return;
  // de minuto em minuto, com a aba visível: recarrega o que muda sozinho (kits, e-mails, admissões) e acende o sino
  SINO_TIMER = setInterval(async () => {
    if (document.hidden || !E.perfil || !$('#main')) return;
    const antes = notificacoes().filter((x) => !vistosSino()[x.chave]).length;
    try {
      const tarefas = [];
      if (tem('admissao', 'signatario') && api.kitsOnline) tarefas.push(api.kitsOnline().then((r) => (E.kitsOnline = r || [])).catch(() => {}));
      if (tem('admissao')) tarefas.push(api.emails().then((r) => (E.emails = r || [])).catch(() => {}));
      if (tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario') || souSeguranca() || souTi()) tarefas.push(api.admissoes().then((r) => (E.admissoes = r || [])).catch(() => {}));
      await Promise.all(tarefas);
    } catch (e) { return; }
    atualizarSino();
    const lista = notificacoes().filter((x) => !vistosSino()[x.chave]);
    if (lista.length > antes) toast((lista.length - antes === 1 ? 'Novidade: ' : plural(lista.length - antes, 'novidade', 'novidades') + ': ') + lista[0].texto);
    if (['kit', 'inicio'].includes(E.view) && !$('.gaveta')) render();
  }, 60000);
}

// Busca de pessoa: nome ou CPF, abre a ficha
function buscar() {
  const q = semAcento($('#busca').value).replace(/[.\-]/g, ''); const box = $('#buscaRes');
  if (q.length < 2) { box.hidden = true; box.innerHTML = ''; return; }
  const achados = E.admissoes.filter((a) => semAcento(a.nome).includes(q) || (/^\d{3,}$/.test(q) && (a.cpf || '').includes(q))).slice(0, 8);
  box.hidden = false;
  box.innerHTML = achados.length
    ? achados.map((a) => `<button data-act="fichaBusca" data-id="${a.id}"><b>${esc(titulo(a.nome))}</b><span>${esc(a.cargo)} · ${esc(a.local)} · ${dmd(a.data_inicio)}</span>${chipEtapa(a)}</button>`).join('')
    : '<div class="busca-vazio">Ninguém com esse nome ou CPF.</div>';
}
function render() {
  if (!E.perfil || !$('#main')) return;
  const itens = navVisivel();
  if (!itens.some((i) => i[0] === E.view)) E.view = itens[0][0];
  const cont = (id) => {
    if (id === 'documentos') return vivas().filter((a) => resumo(a).analise > 0).length;
    if (id === 'remuneracao') return vivas().filter((a) => a.remuneracao === 'pendente' && !a.matricula).length;
    if (id === 'emails') return E.emails.filter((m) => m.status === 'erro').length;
    if (id === 'entregas') return E.entregas.filter((x) => !x.pronto_em).length;
    return 0;
  };
  const item = ([id, n, ic]) => {
    const k = cont(id);
    return `<button data-act="ir" data-v="${id}" ${E.view === id ? 'aria-current="page"' : ''}><span class="ponto"></span>${MI[ic]}<span class="rot-item">${esc(n)}</span>${k ? `<span class="cont num">${k}</span>` : ''}</button>`;
  };
  const sec2 = itens.filter((i) => i[4] === 2);
  $('#menu').innerHTML = `<div class="marca-lateral"><span class="pilula">CIMED</span><div><b>Plataforma de Admissão</b><small>Admissão e onboarding</small></div></div>`
    + itens.filter((i) => i[4] === 1).map(item).join('')
    + (sec2.length ? '<div class="divisor"></div><div class="rot">Configuração</div>' + sec2.map(item).join('') : '')
    + `<div class="rodape-menu"><button data-act="recarregar">Atualizar dados</button><button data-act="tema">Tema claro ou escuro</button></div>`;
  $('#main').innerHTML = (VIEWS[E.view] || VIEWS.inicio)();
  if ($('#esteiraHost')) $('#esteiraHost').hidden = E.view !== 'kit';
  POS[E.view]?.();
}

// =========================================================== CHIPS
function chipEtapa(a) {
  if (a.declinou_em) return chip('c-ag', 'cr', 'Declinou');
  const m = { aguardando: ['c-ag', 'ag'], pendencia: ['c-at', 'at'], conferencia: ['c-info', 'info'], docs_ok: ['c-info', 'info'], registro: ['c-ok', 'ok'], assinatura: ['c-info', 'ag'], pronto: ['c-ok', 'ok'] }[a.etapa] || ['c-ag', ''];
  const falta = a.etapa === 'assinatura' && a.kit_status !== 'concluido' ? faltaAssinar(a) : [];
  return chip(m[0], m[1], ETAPA[a.etapa]?.n || a.etapa) + (falta.length ? ` <span class="apoio" style="font-size:11.5px;white-space:nowrap">falta assinar: ${esc(falta.join(', '))}</span>` : '');
}
function chipDocs(a) { const r = resumo(a); const cls = r.ok === r.total ? ['c-ok', 'ok'] : r.reenviar || r.total - r.ok - r.analise ? ['c-at', 'at'] : ['c-info', 'info']; return chip(cls[0], cls[1], `Docs ${r.ok}/${r.total}`); }
function chipAso(a) { const s = a.aso_status; if (s === 'apto') return chip('c-ok', 'ok', 'ASO apto' + (a.aso_data ? ' ' + dm(a.aso_data) : '')); if (s === 'apto_restricao') return chip('c-ok', 'ok', 'ASO apto com restrição' + (a.aso_data ? ' ' + dm(a.aso_data) : '')); if (s === 'agendado') return chip('c-ag', 'ag', 'ASO ' + (a.aso_data ? dm(a.aso_data) + (a.aso_hora ? ' ' + a.aso_hora : '') : 'agendado') + (a.aso_local ? ' · ' + a.aso_local : '')); if (s === 'inapto') return chip('c-cr', 'cr', 'ASO inapto'); return chip('c-at', 'at', 'ASO a agendar'); }
const asoOk = (a) => a.aso_status === 'apto' || a.aso_status === 'apto_restricao';
// Segurança do Trabalho: a Ordem de Serviço vale para todas as unidades (Josianny, 20/10/2026)
const temOs = (a) => a.os_status !== 'nao_aplica';
const minhaArea = () => (E.verComo ? (E.verComo.area_id || '') : (E.perfil && E.perfil.area_id) || '');
const souSeguranca = () => tem('admissao') || (tem('especialista') && ['seguranca', 'seguranca_prev'].includes(minhaArea()));
const souTi = () => tem('admissao') || (tem('especialista') && ['ti', 'tiprev', 'ti_cancela'].includes(minhaArea()));
function chipOs(a) { if (!temOs(a)) return ''; return a.os_status === 'validada' ? chip('c-ok', 'ok', 'OS validada' + (a.os_validada_em ? ' ' + dm(new Date(a.os_validada_em)) : '')) : (a.os_atividades || a.os_riscos || a.os_epi) ? chip('c-ag', 'ag', 'OS pré-preenchida, falta validar') : chip('c-at', 'at', 'OS sem dados'); }
// TI: o chamado precisa sair com antecedência (prazo_ti_horas do local; 48 h, CDs 72 h), em dias úteis
function prazoChamadoTi(a) {
  if (!(a.ti_itens || []).length || !a.data_inicio) return null;
  const horas = localDe(a.local).prazo_ti_horas || 48; let d = deIso(a.data_inicio); let uteis = Math.ceil(horas / 24);
  while (uteis > 0) { d = new Date(d.getTime() - 86400000); if (d.getDay() !== 0 && d.getDay() !== 6) uteis--; }
  return d;
}
const chamadoTiAtrasado = (a) => { const p = prazoChamadoTi(a); return !!p && !a.chamado_ti && !a.declinou_em && a.kit_status !== 'concluido' && p.getTime() - Date.now() < 86400000; };
function chipRemun(a) { return a.remuneracao === 'ok' ? chip('c-ok', 'ok', 'Remuneração') : a.remuneracao === 'ocupada' ? chip('c-cr', 'cr', 'Posição ocupada') : chip('c-at', 'at', 'Remuneração pendente'); }

function faixaCorte(turma) {
  const corte = corteTurma(turma), falta = corte.getTime() - Date.now();
  const t = E.turmasCal.find((x) => x.data === turma && x.ativo);
  const cal = t ? ` · ${esc(t.nome || GRUPO_TURMA[t.grupo] || '')}${t.corte_inclusao ? ' · inclusão até ' + dmh(t.corte_inclusao) : ''}` : ' · fora do cronograma de turmas';
  const doTurma = vivas().filter((a) => a.data_inicio === turma);
  const pend = doTurma.filter((a) => ['aguardando', 'pendencia', 'conferencia'].includes(a.etapa)).length;
  if (falta <= 0) return `<div class="faixa critica"><span class="relogio">Corte encerrado</span><p>O corte de documentos da turma de ${dmd(turma)} foi ${dmh(corte)}${cal}. ${pend ? plural(pend, 'pessoa ainda tem', 'pessoas ainda têm') + ' documentação incompleta.' : 'Todos estão com a documentação completa.'}</p></div>`;
  const h = Math.floor(falta / 3600e3), mi = Math.floor(falta % 3600e3 / 60e3);
  const cls = falta < 4 * 3600e3 ? 'critica' : falta < DIA ? 'quente' : '';
  const txt = h >= 48 ? Math.floor(h / 24) + ' dias e ' + (h % 24) + ' h' : h + ' h ' + String(mi).padStart(2, '0') + ' min';
  return `<div class="faixa ${cls}"><span class="relogio num">Faltam ${txt}</span><p>para o corte de documentos da turma de <b>${dmd(turma)}</b> (${dmh(corte)}${cal}). ${pend ? plural(pend, 'pessoa com documentação incompleta', 'pessoas com documentação incompleta') + '.' : 'Ninguém com pendência nesta turma.'}</p>${pend && tem('admissao') ? `<button class="bt pri" data-act="cobrarTurma" data-t="${turma}">Cobrar todos com pendência</button>` : ''}</div>`;
}

// ============================================================ TELAS
const VIEWS = {}, POS = {};

// o que precisa de atenção de quem está logado: Início e sino
function listaAtencao() {
  const v = vivas();
  const atencao = [];
  v.forEach((a) => {
    if (risco(a)) atencao.push([0, a, 'Integração ' + dmd(a.data_inicio) + ' e documentação incompleta']);
    if (a.etapa === 'docs_ok' && a.remuneracao !== 'ok') atencao.push([1, a, 'Falta a Remuneração validar a posição ' + (a.posicao || '')]);
    if (a.aso_status === 'inapto' && !a.declinou_em && !a.matricula) atencao.push([0, a, 'ASO inapto: a Saúde decide na ficha se remarca o exame ou cancela a admissão']);
    if (a.etapa === 'docs_ok' && !asoOk(a) && a.aso_status !== 'inapto') atencao.push([1, a, 'Documentação completa, falta o ASO']);
    if (tem('admissao') && chamadoTiAtrasado(a)) atencao.push([1, a, 'Chamado de TI ainda não aberto: o prazo da TI é até ' + dm(prazoChamadoTi(a))]);
    if (souSeguranca() && temOs(a) && a.os_status !== 'validada' && !a.declinou_em && a.etapa !== 'pronto') atencao.push([2, a, 'Ordem de Serviço da Segurança ainda não validada']);
    if (faltaLg(a)) atencao.push([1, a, 'ASO apto: falta a recrutadora movimentar no LG']);
    if (tem('admissao') && faltaDossie(a)) atencao.push([2, a, 'Kit assinado: exportar o dossiê antes do descarte']);
    if (resumo(a).analise) atencao.push([2, a, plural(resumo(a).analise, 'documento recebido para conferir', 'documentos recebidos para conferir')]);
  });
  atencao.sort((x, y) => x[0] - y[0]);
  return atencao;
}
VIEWS.inicio = () => {
  const v = vivas();
  const prox = turmas().find((t) => deIso(t) >= deIso(iso(new Date())));
  const n = (f) => v.filter(f).length;
  const max = Math.max(1, ...ETAPAS.map((e) => n((a) => a.etapa === e.id)));
  const atencao = listaAtencao();
  return `<div class="cab"><div class="t"><span class="olho">Visão geral</span><h1>Admissões em andamento</h1><p class="apoio">Tudo o que vivia na planilha de admissão, atualizado por quem faz cada etapa. A etapa de cada pessoa é calculada pelo sistema a partir dos documentos, do ASO e da remuneração.</p></div>
    <div class="linha-flex">${tem('admissao') ? '<button class="bt" data-act="importar">Importar a planilha atual (.xlsx)</button>' : ''}<button class="bt ama" data-act="ir" data-v="novo">${MI.novo} Cadastrar aprovado</button></div></div>
  ${prox ? faixaCorte(prox) : ''}
  <div class="grade g4">
    <div class="caixa kpi"><b class="num">${v.length}</b><span>em processo de admissão</span></div>
    <div class="caixa kpi"><b class="num">${n((a) => ['aguardando', 'pendencia'].includes(a.etapa))}</b><span>esperando documento do candidato</span></div>
    <div class="caixa kpi"><b class="num">${n((a) => resumo(a).analise > 0)}</b><span>com documento para conferir</span></div>
    <div class="caixa kpi"><b class="num">${n((a) => a.etapa === 'registro')}</b><span>prontos para gerar matrícula</span></div>
  </div>
  <div class="grade g2">
    <div class="caixa pilha"><h2>Onde cada pessoa está</h2><div class="funil">${ETAPAS.map((e) => { const k = n((a) => a.etapa === e.id); return `<div class="l"><span>${esc(e.n)}</span><div class="barra"><i style="width:${k / max * 100}%"></i></div><b class="num">${k}</b></div>`; }).join('')}</div>
      <p class="apoio" style="font-size:12.5px">Declinaram: ${E.admissoes.filter((a) => a.declinou_em).length}.</p></div>
    <div class="caixa pilha"><h2>Precisa de atenção</h2>${atencao.length ? `<div class="tab-wrap"><table><tbody>${atencao.slice(0, 8).map(([p, a, t]) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.cargo)} · ${esc(a.local)}</span></td><td>${p === 0 ? chip('c-cr', 'cr', t) : esc(t)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nada urgente agora.</div>'}</div>
  </div>`;
};

// ------------------------------------------------------------ cadastro
// Linhas de dependentes (cadastro e ficha): nome, parentesco, nascimento, CPF e para que serve
function linhaDep(d, i, pref) {
  const opt = PARENTESCOS.map((p) => `<option ${p === d.parentesco ? 'selected' : ''}>${p}</option>`).join('');
  const chk = (f, rot) => `<label><input type="checkbox" data-dep="${i}" data-f="${f}" ${d[f] ? 'checked' : ''}> ${rot}</label>`;
  return `<div class="dep-linha" data-i="${i}">
    <input data-dep="${i}" data-f="nome" value="${esc(d.nome || '')}" placeholder="Nome completo" aria-label="Nome do dependente">
    <select data-dep="${i}" data-f="parentesco" aria-label="Parentesco"><option value="">Parentesco</option>${opt}</select>
    <input data-dep="${i}" data-f="nasc" type="date" value="${esc(d.nasc || '')}" aria-label="Nascimento">
    <input data-dep="${i}" data-f="cpf" value="${esc(d.cpf || '')}" placeholder="CPF" inputmode="numeric" aria-label="CPF do dependente">
    <div class="marcas">${chk('plano_saude', 'Plano de saúde')}${chk('odonto', 'Odonto')}${chk('salario_familia', 'Salário-família')}${chk('universitario', 'Universitário(a)')}</div>
    <button type="button" class="bt peq fan" data-act="${pref}DepDel" data-i="${i}" ${d.id ? `data-id="${d.id}"` : ''}>Remover</button>
  </div>`;
}
function blocoDeps(lista, pref, podeEditar) {
  if (!ligado('dependentes_ativo')) return '';
  return `<div class="deps" id="${pref}Deps">${lista.map((d, i) => (podeEditar ? linhaDep(d, i, pref) : `<div class="dep-linha so-leitura"><b>${esc(titulo(d.nome))}</b><span>${esc(d.parentesco)}${d.idade != null ? ' · ' + d.idade + ' anos' : ''}</span><span class="apoio">${[d.plano_saude && 'plano', d.odonto && 'odonto', d.salario_familia && 'salário-família', d.universitario && 'universitário(a)'].filter(Boolean).join(' · ')}</span></div>`)).join('') || '<span class="apoio" style="font-size:12.5px">Nenhum dependente informado.</span>'}</div>
  ${podeEditar ? `<div class="linha-flex"><button type="button" class="bt peq" data-act="${pref}DepAdd">+ Adicionar dependente</button><span class="ajuda">Filho(a), enteado(a), cônjuge, companheiro(a), pai ou mãe. Cada um gera a própria lista de documentos (certidão, CPF, SUS, vacinação até 6 anos ou escola acima de 6).</span></div>` : ''}`;
}
// turmas do cronograma que servem para um local (grupo do local ou todas), de hoje em diante
function turmasParaLocal(local) {
  const g = localDe(local).grupo_turma || 'pa_sp_comercial'; const hoje = iso(new Date());
  return E.turmasCal.filter((t) => t.ativo && t.data >= hoje && (t.grupo === 'todos' || t.grupo === g)).sort((a, b) => a.data.localeCompare(b.data));
}
function opcoesTurma(local, sel) {
  const ts = turmasParaLocal(local);
  const op = ts.map((t) => { const tarde = Date.now() > new Date(t.corte_inclusao).getTime(); return `<option value="${t.id}" ${t.id === sel ? 'selected' : ''}>${dmd(t.data)} · ${esc(t.nome || GRUPO_TURMA[t.grupo])} · inclusão até ${dmh(t.corte_inclusao)}${tarde ? ' (passou)' : ''}</option>`; });
  if (!ts.length) { const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); for (let i = 0; i < 6; i++) { const x = new Date(d.getTime() + i * 7 * DIA); op.push(`<option value="data:${iso(x)}">${dmd(x)} · corte ${dmd(corteDe(iso(x)))} 12h (sem cronograma)</option>`); } }
  op.push(`<option value="outra" ${sel === 'outra' ? 'selected' : ''}>Outra data, fora do cronograma</option>`);
  return op.join('');
}
VIEWS.novo = () => {
  const opt = (arr, sel) => arr.map((v) => `<option ${v === sel ? 'selected' : ''}>${esc(v)}</option>`).join('');
  const deps = ligado('dependentes_ativo');
  return `<div class="cab"><div class="t"><span class="olho">Atração e Seleção</span><h1>Cadastrar aprovado</h1><p class="apoio">Ao salvar, a Saúde, a Admissão, a Remuneração e Benefícios recebem o aviso, e o candidato recebe a lista de documentos do perfil dele com o prazo do corte da turma. Quem tem acesso acompanha tudo pelo painel.</p></div></div>
  <form class="caixa" id="fNovo" novalidate><div class="form">
    <div class="secao-form">Candidato</div>
    <div class="campo l2"><label for="n_nome">Nome completo, como no documento</label><input id="n_nome" autocomplete="off"></div>
    <div class="campo"><label for="n_cpf">CPF</label><input id="n_cpf" inputmode="numeric" placeholder="000.000.000-00"></div>
    <div class="campo"><label for="n_nasc">Data de nascimento</label><input id="n_nasc" type="date"></div>
    <div class="campo"><label for="n_tel">WhatsApp</label><input id="n_tel" inputmode="tel" placeholder="(35) 99999-9999"></div>
    <div class="campo l2"><label for="n_email">E-mail pessoal</label><input id="n_email" type="email"><span class="ajuda">Obrigatório: é por onde vão a lista de documentos, as cobranças e o convite da integração.</span></div>
    <div class="campo"><label for="n_sexo">Sexo</label><select id="n_sexo"><option value="F">Feminino</option><option value="M">Masculino</option></select><span class="ajuda">Define se pede reservista.</span></div>
    <div class="campo"><label for="n_civil">Estado civil</label><select id="n_civil">${opt(['Solteiro(a)', 'Casado(a)', 'União estável', 'Divorciado(a)', 'Viúvo(a)'])}</select><span class="ajuda">Casado(a) ou união estável: certidão, RG e SUS do cônjuge. Divorciado(a): certidão com averbação.</span></div>
    <div class="campo"><label for="n_filhos">Filhos</label><input id="n_filhos" type="number" min="0" max="15" value="0" ${deps ? 'readonly title="Vem da lista de dependentes"' : ''}>${deps ? '<span class="ajuda">Contado pela lista de dependentes abaixo.</span>' : ''}</div>
    <div class="campo"><label for="n_tipo">Contratação</label><select id="n_tipo">${opt(['CLT', 'Jovem aprendiz', 'Estágio', 'Temporário'])}</select></div>
    <div class="campo"><label for="n_agencia">Agência</label><input id="n_agencia" placeholder="Só para temporário"><span class="ajuda">Temporário é contratado pela agência</span></div>
    <div class="campo"><label for="n_pcd">PCD</label><select id="n_pcd">${opt(['Não', 'Auditiva', 'Física', 'Intelectual', 'Visual', 'Mental', 'Reabilitado'])}</select></div>
    <div class="campo"><label for="n_estr">Nacionalidade</label><select id="n_estr"><option value="0">Brasileira</option><option value="1">Estrangeira</option></select></div>
    ${deps ? `<div class="secao-form">Dependentes</div><div class="campo l4">${blocoDeps(E.depsNovo, 'n', true)}</div>` : ''}
    <div class="secao-form">Vaga</div>
    <div class="campo"><label for="n_cargo">Cargo</label><input id="n_cargo"></div>
    <div class="campo"><label for="n_cbo">CBO</label><input id="n_cbo" inputmode="numeric" placeholder="0000-00"><span class="ajuda">Vai na Ordem de Serviço da Segurança.</span></div>
    <div class="campo"><label for="n_setor">Setor</label><input id="n_setor"></div>
    <div class="campo"><label for="n_local">Local de trabalho</label><select id="n_local">${E.locais.map((l) => `<option ${l.nome === 'Pouso Alegre' ? 'selected' : ''}>${esc(l.nome)}</option>`).join('')}</select><span class="ajuda" id="n_localAj"></span></div>
    <div class="campo"><label for="n_turma">Turma de integração</label><select id="n_turma">${opcoesTurma('Pouso Alegre')}</select><span class="ajuda" id="n_turmaAj"></span></div>
    <div class="campo" id="n_inicioBox" hidden><label for="n_inicio">Data de início</label><input id="n_inicio" type="date"><span class="ajuda">Fora do cronograma: o DP recebe um alerta.</span></div>
    <div class="campo"><label for="n_cc_cod">Centro de custo (código)</label><input id="n_cc_cod" inputmode="numeric"></div>
    <div class="campo"><label for="n_cc_nome">Centro de custo (nome)</label><input id="n_cc_nome"></div>
    <div class="campo l2"><label for="n_escala">Escala</label><input id="n_escala" placeholder="05:30 11:00/12:00 15:18"></div>
    <div class="campo"><label for="n_turno">Turno</label><select id="n_turno"><option value="">—</option>${opt(['1º turno', '2º turno', '3º turno', 'Administrativo', 'Comercial'])}</select></div>
    <div class="campo"><label for="n_posicao">Posição</label><input id="n_posicao" inputmode="numeric"></div>
    <div class="campo"><label for="n_proc">Nº do processo seletivo</label><input id="n_proc" inputmode="numeric"></div>
    <div class="campo l2"><label for="n_req">Requisitante (gestor)</label><input id="n_req"></div>
    <div class="campo l2"><label for="n_req_email">E-mail do requisitante</label><input id="n_req_email" type="email" placeholder="nome@grupocimed.com.br"><span class="ajuda">É para onde vão o aviso ao gestor e a cópia do convite.</span></div>
    <div class="campo l2"><label>Vaga</label><div class="marcas"><label><input type="checkbox" id="n_subst"> Substituição</label><label><input type="checkbox" id="n_nr12"> Integração NR-12 (opera máquina)</label></div></div>
    <div class="campo l2" id="n_substBox" hidden><label for="n_subst_nome">Pessoa substituída</label><input id="n_subst_nome"></div>
    <div class="secao-form">Integração e áreas</div>
    <div class="campo"><label for="n_formato">Formato</label><select id="n_formato"><option value="">Padrão do local</option><option value="presencial">Presencial</option><option value="online">On-line (Teams)</option></select></div>
    <div class="campo l2"><label for="n_link">Link do Teams ou endereço</label><input id="n_link" placeholder="Padrão da turma, se houver"></div>
    <div class="campo"><label for="n_contato">Quem procurar ao chegar</label><input id="n_contato" placeholder="Nome e ramal"></div>
    <div class="campo"><label for="n_camisa">Tamanho de camisa</label><select id="n_camisa">${opt(['PP', 'P', 'M', 'G', 'GG', 'XGG', 'Não usa uniforme'], 'M')}</select></div>
    <div class="campo"><label for="n_fretado">Linha do fretado</label><input id="n_fretado" placeholder="Não usa"></div>
    <div class="campo"><label for="n_cd_entrega">CD de entrega do equipamento</label><input id="n_cd_entrega" placeholder="Só para o comercial"><span class="ajuda">A TI não entrega em casa: o equipamento vai para o CD mais perto do colaborador.</span></div>
    <div class="campo l2"><label>Usa veículo da frota?</label><div class="marcas"><label><input type="checkbox" id="n_frota"> Sim, vai dirigir carro da empresa</label><label><input type="checkbox" id="n_motorista"> Motorista profissional (dirigir é a função)</label></div></div>
    <div class="campo l4"><label>Equipamentos e acessos</label><div class="marcas" id="n_ti">${['Notebook', 'Desktop', 'Usuário de rede', 'E-mail', 'Teams', 'SAP', 'Celular'].map((x) => `<label><input type="checkbox" value="${x}"> ${x}</label>`).join('')}</div><span class="ajuda">A TI recebe a previsão agora e o pedido com a matrícula.</span></div>
  </div><div class="linha-flex" style="margin-top:18px;justify-content:flex-end"><span class="apoio" id="n_resumo" style="flex:1;min-width:220px"></span><button type="submit" class="bt pri" id="n_salvar">Salvar e avisar as áreas</button></div></form>`;
};
POS.novo = () => {
  const f = $('#fNovo'); let nr12Manual = false;
  const upd = () => {
    const l = localDe($('#n_local').value);
    $('#n_localAj').textContent = l.nome ? `Integração: ${l.integ}, ${l.hora}, ${plural(l.dias, 'dia', 'dias')} · calendário ${GRUPO_TURMA[l.grupo_turma] || ''}` : '';
    const t = E.turmasCal.find((x) => x.id === $('#n_turma').value);
    $('#n_inicioBox').hidden = $('#n_turma').value !== 'outra';
    $('#n_turmaAj').textContent = t ? (Date.now() > new Date(t.corte_inclusao).getTime() ? 'O corte de inclusão desta turma já passou: combine com o DP antes de salvar.' : `Documentos até ${dmh(t.corte_documentos)}${t.corte_remuneracao ? ' · Remuneração até ' + dmh(t.corte_remuneracao) : ''}`) : '';
    $('#n_substBox').hidden = !$('#n_subst').checked;
    if (!nr12Manual) $('#n_nr12').checked = l.tipo === 'fabrica' && /operador|produc|produç|máquina|maquina/i.test($('#n_cargo').value);
    if (ligado('dependentes_ativo')) $('#n_filhos').value = E.depsNovo.filter((d) => ['Filho(a)', 'Enteado(a)'].includes(d.parentesco)).length;
  };
  $('#n_local').addEventListener('change', () => {
    $('#n_turma').innerHTML = opcoesTurma($('#n_local').value);
    // comercial: notebook e celular sempre (reunião com a TI, 15/10/2026)
    if (localDe($('#n_local').value).tipo === 'comercial') $('#n_ti').querySelectorAll('input').forEach((c) => { if (['Notebook', 'Celular', 'E-mail', 'Teams'].includes(c.value)) c.checked = true; });
    upd();
  });
  $('#n_nr12').addEventListener('change', () => { nr12Manual = true; });
  f.addEventListener('change', upd); f.addEventListener('input', (e) => { if (e.target.id === 'n_cargo') upd(); });
  f.addEventListener('input', (e) => { const el = e.target; if (el.dataset.dep == null) return; const d = E.depsNovo[+el.dataset.dep]; if (!d) return; d[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value; upd(); });
  upd();
  f.addEventListener('submit', async (e) => {
    e.preventDefault();
    f.querySelectorAll('.erro').forEach((x) => x.remove());
    const v = (id) => $('#' + id).value.trim(); const erros = [];
    const marca = (id, m) => { const s = document.createElement('span'); s.className = 'erro'; s.textContent = m; $('#' + id).parentElement.appendChild(s); erros.push(id); };
    if (v('n_nome').split(/\s+/).length < 2) marca('n_nome', 'Digite o nome completo, como está no documento.');
    if (!cpfValido(v('n_cpf'))) marca('n_cpf', 'CPF inválido. Confira os números.');
    if (v('n_tipo') === 'Temporário' && !v('n_agencia')) marca('n_agencia', 'Informe a agência que contrata o temporário.');
    if (!v('n_nasc')) marca('n_nasc', 'Informe a data de nascimento.');
    if (!/^\S+@\S+\.\S+$/.test(v('n_email'))) marca('n_email', 'E-mail inválido.');
    if (v('n_req_email') && !/^\S+@\S+\.\S+$/.test(v('n_req_email'))) marca('n_req_email', 'E-mail inválido.');
    if (v('n_tel').replace(/\D/g, '').length < 10) marca('n_tel', 'Informe o WhatsApp com DDD.');
    ['n_cargo', 'n_req'].forEach((id) => { if (!v(id)) marca(id, 'Obrigatório.'); });
    const turmaSel = v('n_turma'); let turma_id = null, data_inicio = null;
    if (turmaSel === 'outra') { if (!v('n_inicio')) marca('n_inicio', 'Informe a data de início.'); data_inicio = v('n_inicio'); }
    else if (turmaSel.startsWith('data:')) data_inicio = turmaSel.slice(5);
    else { const t = E.turmasCal.find((x) => x.id === turmaSel); turma_id = t?.id || null; data_inicio = t?.data || null; if (!data_inicio) marca('n_turma', 'Escolha a turma.'); }
    const deps = E.depsNovo.filter((d) => (d.nome || '').trim() || d.parentesco);
    deps.forEach((d, i) => {
      if ((d.nome || '').trim().split(/\s+/).length < 2 || !d.parentesco) { erros.push('n_turma'); toast(`Dependente ${i + 1}: nome completo e parentesco são obrigatórios.`, true); }
      if (d.cpf && !cpfValido(d.cpf)) { erros.push('n_turma'); toast(`Dependente ${i + 1}: CPF inválido.`, true); }
    });
    if (erros.length) { $('#' + erros[0])?.focus(); return; }
    $('#n_salvar').disabled = true;
    const novo = await tentar(() => api.criar({
      nome: v('n_nome'), cpf: v('n_cpf').replace(/\D/g, ''), nasc: v('n_nasc'), email: v('n_email'), telefone: v('n_tel'),
      sexo: v('n_sexo'), estado_civil: v('n_civil'), filhos: +v('n_filhos') || 0, tipo_contrato: v('n_tipo'), pcd: v('n_pcd'),
      estrangeiro: v('n_estr') === '1', cargo: v('n_cargo'), setor: v('n_setor') || null, local: v('n_local'), data_inicio, turma_id,
      escala: v('n_escala') || null, posicao: v('n_posicao') || null, processo: v('n_proc') || null, requisitante: v('n_req'),
      requisitante_email: v('n_req_email').toLowerCase() || null, cc_codigo: v('n_cc_cod') || null, cc_nome: v('n_cc_nome') || null,
      cbo: v('n_cbo') || null, cd_entrega: v('n_cd_entrega') || null,
      turno: v('n_turno') || null, nr12: $('#n_nr12').checked, substituicao: $('#n_subst').checked, pessoa_substituida: v('n_subst_nome') || null,
      formato_integ: v('n_formato') || null, link_ou_endereco: v('n_link') || null, contato_chegada: v('n_contato') || null,
      camisa: v('n_camisa'), fretado: v('n_fretado') || null, usa_frota: $('#n_frota').checked, motorista_profissional: $('#n_motorista').checked, agencia: v('n_agencia') || null,
      ti_itens: [...document.querySelectorAll('#n_ti input:checked')].map((x) => x.value),
    }));
    if (novo && deps.length) {
      for (const d of deps) await tentar(() => api.salvarDependente({ admissao_id: novo.id, nome: d.nome.trim(), parentesco: d.parentesco, nasc: d.nasc || null, cpf: (d.cpf || '').replace(/\D/g, '') || null, universitario: !!d.universitario, plano_saude: !!d.plano_saude, odonto: !!d.odonto, salario_familia: !!d.salario_familia }));
    }
    $('#n_salvar') && ($('#n_salvar').disabled = false);
    if (novo) {
      E.depsNovo = [];
      toast('Aprovado cadastrado. ' + (avisadas('cadastro', novo) || 'Nenhuma área com e-mail para avisar ainda.') + ' A lista de documentos foi para a fila do candidato.');
      E.view = 'painel'; E.turma = novo.data_inicio; await recarregar(); abrirFicha(novo.id);
    }
  });
};

// -------------------------------------------------------------- painel
function filtradas() { return E.admissoes.filter((a) => (E.turma === 'todas' || a.data_inicio === E.turma) && (!E.localF || a.integ === E.localF)); }
VIEWS.painel = () => {
  const ts = turmas(); const locais = [...new Set(E.admissoes.map((a) => a.integ).filter(Boolean))].sort();
  const lista = filtradas();
  return `<div class="cab"><div class="t"><span class="olho">Admissão</span><h1>Painel da admissão</h1><p class="apoio">A planilha de admissão em forma de quadro. Cada área atualiza a sua parte e nenhuma coluna aceita texto livre: as observações ficam na ficha, com autor e data.</p></div>
    ${tem('admissao') ? '<button class="bt" data-act="importar">Importar a planilha atual (.xlsx)</button>' : ''}</div>
  ${E.turma !== 'todas' && E.turma ? faixaCorte(E.turma) : ''}
  <div class="filtros"><label for="fTurma" class="olho">Turma</label><select id="fTurma"><option value="todas">Todas as turmas</option>${ts.map((t) => `<option value="${t}" ${t === E.turma ? 'selected' : ''}>${dmd(t)}</option>`).join('')}</select>
    <label for="fLocal" class="olho">Integração</label><select id="fLocal"><option value="">Todos os locais</option>${locais.map((l) => `<option ${l === E.localF ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>
    <span style="flex:1"></span><div class="seg" role="group" aria-label="Modo"><button data-act="modo" data-m="quadro" aria-pressed="${E.modo === 'quadro'}">Quadro</button><button data-act="modo" data-m="tabela" aria-pressed="${E.modo === 'tabela'}">Tabela</button></div></div>
  ${!E.admissoes.length ? `<div class="vazio">Nenhuma admissão ainda. ${tem('rs', 'admissao') ? 'Cadastre um aprovado ou importe a planilha atual.' : ''}</div>` : E.modo === 'quadro' ? quadro(lista) : tabela(lista)}`;
};
POS.painel = () => {
  $('#fTurma')?.addEventListener('change', (e) => { E.turma = e.target.value; render(); });
  $('#fLocal')?.addEventListener('change', (e) => { E.localF = e.target.value; render(); });
};
function cartao(a) {
  const r = resumo(a); const chips = [];
  if (['aguardando', 'pendencia', 'conferencia'].includes(a.etapa)) {
    (r.pendentes || []).slice(0, 3).forEach((p) => chips.push(chip('c-ag', 'ag', String(p).replace(', frente e verso', ''))));
    if ((r.pendentes || []).length > 3) chips.push(chip('c-ag', '', '+' + (r.pendentes.length - 3)));
    if (r.analise) chips.push(chip('c-info', 'info', r.analise + ' para conferir'));
  }
  if (['docs_ok', 'registro'].includes(a.etapa)) { chips.push(chipAso(a)); chips.push(chipRemun(a)); }
  if (['assinatura', 'pronto'].includes(a.etapa)) chips.push(chip('c-esc', '', a.matricula || ''));
  if (foraCorte(a)) chips.unshift(chip('c-cr', 'cr', 'Fora do corte'));
  return `<button class="cartao ${risco(a) ? 'risco' : ''}" data-act="ficha" data-id="${a.id}"><span class="nm">${esc(titulo(a.nome))}</span><span class="sub">${esc(a.cargo)} · ${esc(a.local)}</span><span class="chips">${chips.join('')}</span><span class="sub num">Integração ${dmd(a.data_inicio)}${a.cobrancas ? ' · ' + plural(a.cobrancas, 'cobrança', 'cobranças') : ''}</span></button>`;
}
function quadro(lista) {
  const decl = lista.filter((a) => a.declinou_em).length;
  return `<div class="quadro-wrap"><div class="quadro">${ETAPAS.map((e) => { const cs = lista.filter((a) => !a.declinou_em && a.etapa === e.id); return `<section class="coluna" aria-label="${esc(e.n)}"><h3>${esc(e.n)}<span class="n num">${cs.length}</span></h3><p class="dica">${esc(e.d)}</p>${cs.map(cartao).join('') || '<div class="vazio-col">Ninguém nesta etapa</div>'}</section>`; }).join('')}</div></div>
  ${decl ? `<p class="apoio" style="font-size:12.5px">${plural(decl, 'pessoa declinou', 'pessoas declinaram')} nesta seleção. Veja na tabela.</p>` : ''}`;
}
function tabela(lista) {
  return `<div class="tab-wrap"><table><thead><tr><th>Nome</th><th>Cargo · setor</th><th>Integração</th><th>R&amp;S</th><th>Etapa</th><th>Documentos</th><th>ASO</th><th>Remuneração</th><th>Matrícula</th><th>DP</th><th>TI</th><th>Camisa</th><th>Fretado</th></tr></thead><tbody>
  ${lista.map((a) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b></td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.setor || '')}</span></td><td class="num">${dmd(a.data_inicio)}<br><span class="apoio" style="font-size:12px">${esc(a.integ || '')} ${esc(a.hora_integ || '')}</span></td><td>${esc(a.analista_rs || '')}</td><td>${chipEtapa(a)}</td><td>${a.declinou_em ? '—' : chipDocs(a)}</td><td>${a.declinou_em ? '—' : chipAso(a)}</td><td>${a.declinou_em ? '—' : chipRemun(a)}</td><td class="num">${esc(a.matricula || '—')}</td><td>${esc(a.analista_dp || '—')}</td><td>${(a.ti_itens || []).length ? esc(a.ti_itens.join(', ')) : '—'}</td><td>${esc(a.camisa || '—')}</td><td>${esc(a.fretado || '—')}</td></tr>`).join('')}
  </tbody></table></div>`;
}

// ---------------------------------------------------------------- ficha
async function abrirFicha(id) {
  const a = adm(id); if (!a) return;
  $('#camada').innerHTML = `<div class="veu" data-act="fechar"><aside class="gaveta" data-pare="1"><div class="carregando">Abrindo a ficha…</div></aside></div>`;
  const [docs, evs, arqs, acs, deps, opc, termo] = await Promise.all([
    tem('rs', 'admissao', 'saude') ? tentar(() => api.docs(id)) : [],
    tentar(() => api.eventos(id)),
    tem('admissao', 'saude') ? tentar(() => api.arquivos(id)) : [],
    tem('admissao') && api.acessos ? tentar(() => api.acessos(id)) : [],
    ligado('dependentes_ativo') && tem('rs', 'admissao', 'saude', 'remuneracao') ? tentar(() => api.dependentes(id)) : [],
    tem('admissao') ? tentar(() => api.opcoesBeneficios(id)) : null,
    tem('rs', 'admissao') ? tentar(() => api.termoRecebido(id)) : null,
  ]);
  E.acessosFicha = acs || [];
  E.fichaAberta = id; E.docsFicha = docs || []; E.depsFicha = (deps || []).map((d) => Object.assign({}, d)); E.opcFicha = opc || null; E.termoFicha = termo || null;
  desenharFicha(a, docs || [], evs || [], arqs || []);
}
const ST_ARQ = {
  analisando: ['c-info', 'ag', 'IA lendo…'], aceito: ['c-ok', 'ok', 'IA aceitou'], duvida: ['c-info', 'info', 'Conferir'],
  recusado: ['c-at', 'at', 'IA recusou'], erro: ['c-cr', 'cr', 'Leitura falhou'],
};
function listaArquivos(a, tipo, arqs, opc = { enviar: true, remover: true }) {
  const meus = arqs.filter((x) => x.tipo === tipo);
  const lado = (l) => (l === 'frente' || l === 'verso' ? ' · ' + l : l === 'ambos' ? ' · frente e verso' : '');
  return `<div class="pilha" style="gap:6px;margin-top:8px">${meus.map((x) => {
    const s = ST_ARQ[x.status] || ST_ARQ.erro;
    return `<div class="linha-flex" style="gap:8px;font-size:12.5px;align-items:flex-start"><span>${chip(s[0], s[1], s[2] + lado(x.lado))}</span>
      <span style="flex:1;min-width:160px"><b>${esc(x.nome || 'arquivo')}</b>${x.motivo ? ' · ' + esc(x.motivo) : ''}<br><span class="apoio">${esc(x.enviado_por || '')} · ${quando(x.criado_em)}</span></span>
      <button class="bt peq" data-act="verArq" data-id="${x.id}" data-c="${esc(x.caminho)}">Ver</button>${opc.remover ? `<button class="bt peq fan" data-act="tirarArq" data-id="${x.id}" data-c="${esc(x.caminho)}" data-adm="${a.id}">Remover</button>` : ''}</div>`;
  }).join('')}
  ${!a.declinou_em && opc.enviar ? `<div><button class="bt peq" data-act="enviarArq" data-adm="${a.id}" data-tipo="${esc(tipo)}">${meus.length ? 'Enviar outro arquivo' : 'Enviar arquivo'}</button></div>` : ''}</div>`;
}
// resumo das escolhas do kit (opcoes_beneficios.dados), com os rótulos dos formulários da Esteira
function resumoOpcoes(dados) {
  const rot = { unimed_adesao: 'Plano Unimed', unimed_dep: 'Dependentes Unimed', sulamerica_adesao: 'Sul América', sulamerica_dep: 'Dependentes Sul América', odonto_adesao: 'Odontológico', odonto_dep: 'Dependentes odonto', vt_adesao: 'Vale transporte', vt_meio: 'Transporte', vt_qtd: 'Passagens/dia', vt_valor: 'Valor da passagem', fretado_ponto: 'Ponto do fretado', benef_tem: 'Beneficiários do seguro', benef: 'Beneficiários', ir_tem: 'Dependentes IR', ir_dep: 'Dependentes IR', sf_confere: 'Salário-família', sf_correcao: 'Correção', sf_vacina: 'Vacinação', sf_escola: 'Escola', rest_adesao: 'Restaurante', va_adesao: 'VA', hiper_adesao: 'Hiperlife', seguro_adesao: 'Seguro de vida', proposta_adesao: 'Proposta de adesão', linha_adesao: 'Linha móvel', linha_numero: 'Nº da linha', linha_operadora: 'Operadora' };
  return Object.entries(dados || {}).filter(([k, v]) => !/^q\d|^termo|^codigo_conduta/.test(k) && v != null && v !== '' && !(Array.isArray(v) && !v.length)).map(([k, v]) => `<div><span class="apoio">${esc(rot[k] || k)}:</span> <b>${Array.isArray(v) ? esc(v.filter((x) => x && x.nome).map((x) => [x.nome, x.parentesco, x.nasc, x.pct && x.pct + '%'].filter(Boolean).join(' · ')).join('; ') || '—') : esc(String(v))}</b></div>`).join('');
}
function desenharFicha(a, docs, evs, arqs = []) {
  const ehLaudo = (d) => String(d.tipo).split('#')[0] === 'laudo';
  const r = resumo(a); const dp = tem('admissao'); const l = localDe(a.local);
  const proxTurma = turmasParaLocal(a.local).find((t) => t.data > a.data_inicio);
  const prox = proxTurma ? proxTurma.data : iso(new Date(deIso(a.data_inicio).getTime() + 7 * DIA));
  const podeAso = tem('admissao', 'saude'); const podeRem = tem('admissao', 'remuneracao'); const podeRs = tem('rs', 'admissao');
  const exigeLg = ligado('exige_lg_movimentado');
  const selDoc = (d) => `<select class="sel-doc" data-doc="${esc(d.tipo)}" aria-label="Situação de ${esc(d.nome)}">${Object.entries(ST_DOC).map(([k, n]) => `<option value="${k}" ${k === d.status ? 'selected' : ''}>${n}</option>`).join('')}</select>`;
  const stDoc = { ok: chip('c-ok', 'ok', 'Conferido'), analise: chip('c-info', 'info', 'Recebido'), reenviar: chip('c-at', 'at', 'Reenviar'), falta: chip('c-ag', 'ag', 'Não enviado') };
  const turmaCal = E.turmasCal.find((t) => t.id === a.turma_id);
  const chipLg = a.lg_movimentado_em ? chip('c-ok', 'ok', 'Movimentado no LG ' + dm(new Date(a.lg_movimentado_em))) : asoOk(a) ? chip('c-at', 'at', 'Falta movimentar no LG') : chip('c-ag', 'ag', 'Depois do ASO apto');
  const inputFicha = (id, ph, extra = '') => `<input id="${id}" class="sel-doc" placeholder="${esc(ph)}" ${extra}>`;
  $('#camada').innerHTML = `<div class="veu" data-act="fechar"><aside class="gaveta" role="dialog" aria-modal="true" aria-label="Ficha" data-pare="1">
   <div class="linha-flex" style="align-items:flex-start"><div class="pilha" style="flex:1;gap:6px"><span class="olho">Ficha do candidato</span><h2 style="font-size:24px">${esc(titulo(a.nome))}</h2><span class="apoio">${esc(a.cargo)}${a.setor ? ' · ' + esc(a.setor) : ''} · ${esc(a.local)}${a.cc_codigo ? ' · CC ' + esc(a.cc_codigo) : ''}</span>
    <div class="linha-flex">${chipEtapa(a)}${foraCorte(a) ? chip('c-cr', 'cr', 'Fora do corte') : ''}${a.pcd !== 'Não' ? chip('c-info', 'info', 'PCD') : ''}${a.tipo_contrato !== 'CLT' ? chip('c-info', 'info', a.tipo_contrato + (a.agencia ? ' · ' + a.agencia : '')) : ''}${a.nasc && idade(a.nasc) < 18 ? chip('c-at', 'at', 'Menor de idade' + (l.tipo === 'fabrica' ? ' em fábrica' : '')) : ''}${a.motorista_profissional ? chip('c-info', 'info', 'Motorista profissional') : ''}${a.formato_integ === 'online' ? chip('c-ag', '', 'Integração on-line') : ''}${a.nr12 ? chip('c-ag', '', 'NR-12') : ''}${a.arquivado_em ? chip('c-ok', 'ok', 'Dossiê arquivado') : ''}${a.origem === 'importacao' ? chip('c-ag', '', 'Veio da planilha') : ''}</div></div>
    <button class="bt" data-act="fechar">Fechar</button></div>
   ${a.declinou_em ? `<div class="faixa"><p><b>Declinou</b> em ${quando(a.declinou_em)}: ${esc(a.declinou_motivo || '')}.${a.declinou_sensivel ? ' <span class="apoio">Motivo sensível: as áreas receberam só o aviso de cancelamento.</span>' : ''}</p></div>` : ''}
   <div class="grade g2">
    <div class="caixa pilha" style="gap:8px"><h3>Integração</h3><span class="num">${dmd(a.data_inicio)} às ${esc(a.hora_integ || '')}</span><span class="apoio">${esc(a.integ || '')} · ${plural(a.dias_integ || 1, 'dia', 'dias')} · ${a.formato_integ === 'online' ? 'on-line (Teams)' : 'presencial'}${a.nr12 ? ' · NR-12' : ''}</span>
     ${a.link_ou_endereco ? `<span class="apoio">${a.formato_integ === 'online' ? 'Link' : 'Endereço'}: ${esc(a.link_ou_endereco)}</span>` : ''}${a.contato_chegada ? `<span class="apoio">Ao chegar, procurar: ${esc(a.contato_chegada)}</span>` : ''}
     <span class="apoio">Turma: ${turmaCal ? esc(turmaCal.nome || GRUPO_TURMA[turmaCal.grupo]) + ' · inclusão até ' + dmh(turmaCal.corte_inclusao) : 'fora do cronograma'}</span><span class="apoio">Corte de documentos: ${dmh(corteA(a))}</span>
     ${podeRs && !a.declinou_em && !a.matricula ? `<button class="bt peq" data-act="turma" data-id="${a.id}" data-t="${prox}">Mover para a turma de ${dm(prox)}</button>` : ''}</div>
    <div class="caixa pilha" style="gap:8px"><h3>Contato</h3><span>${esc(a.telefone || '—')}</span><span>${esc(a.email || 'sem e-mail')}</span>${tem('saude', 'admissao', 'remuneracao') && a.cpf ? `<span class="apoio num">CPF ${fmtCpf(a.cpf)}${a.nasc ? ' · ' + idade(a.nasc) + ' anos' : ''}</span>` : ''}<span class="apoio">Gestor: ${esc(a.requisitante || '—')}${a.requisitante_email ? ' · ' + esc(a.requisitante_email) : ' · <b>sem e-mail</b> (o aviso ao gestor não sai)'}</span></div>
   </div>
   <div class="caixa pilha"><h3>Checklist da admissão</h3><div class="docs">
    <div class="doc"><div><div class="dn">Exame admissional (ASO)</div><div class="dm">${chipAso(a)}${a.aso_restricao ? ` <span class="apoio" style="font-size:12px">Restrição: ${esc(a.aso_restricao)}</span>` : ''}${a.aso_status === 'inapto' && !a.declinou_em ? '<div class="motivo">Inapto: a Saúde decide aqui se o exame é refeito (preencha data, hora e local e clique em Remarcar) ou se a admissão é cancelada. A matrícula fica bloqueada. R&S e Admissão já foram avisados.</div>' : ''}</div></div><div class="acoes">${podeAso && !a.declinou_em ? `<select class="sel-doc" id="asoSel"><option value="aguardando" ${a.aso_status === 'aguardando' ? 'selected' : ''}>A agendar</option><option value="agendado" ${a.aso_status === 'agendado' ? 'selected' : ''}>Agendado</option><option value="apto" ${a.aso_status === 'apto' ? 'selected' : ''}>Apto</option><option value="apto_restricao" ${a.aso_status === 'apto_restricao' ? 'selected' : ''}>Apto com restrição</option><option value="inapto" ${a.aso_status === 'inapto' ? 'selected' : ''}>Inapto</option></select><input type="date" id="asoData" class="sel-doc" value="${a.aso_data || ''}" aria-label="Data do ASO"><input id="asoHora" class="sel-doc" style="width:70px" value="${esc(a.aso_hora || '')}" placeholder="hora" aria-label="Hora do exame"><input id="asoLocal" class="sel-doc" style="width:150px" value="${esc(a.aso_local || '')}" placeholder="ambulatório ou clínica" aria-label="Local do exame"><input id="asoRestr" class="sel-doc" style="width:150px" value="${esc(a.aso_restricao || '')}" placeholder="restrição, se houver" aria-label="Restrição"><button class="bt peq" data-act="salvarAso" data-id="${a.id}">Salvar</button>${a.aso_status === 'inapto' ? `<button class="bt peq pri" data-act="asoRemarcar" data-id="${a.id}" title="Exame refeito: volta para agendado com a data informada">Remarcar exame</button><button class="bt peq fan" data-act="asoCancelar" data-id="${a.id}" title="Cancela a admissão por ASO inapto; as áreas recebem só o aviso de cancelamento">Cancelar admissão</button>` : ''}` : ''}</div></div>
    ${temOs(a) ? `<div class="doc"><div><div class="dn">Ordem de Serviço · Segurança do Trabalho</div><div class="dm">${chipOs(a)}${a.cbo ? ' ' + chip('c-ag', '', 'CBO ' + a.cbo) : ''}${a.os_validada_por ? ` <span class="apoio" style="font-size:12px">validada por ${esc(a.os_validada_por)}</span>` : ''}${a.os_status !== 'validada' ? '<div class="motivo">A Esteira só gera a OS no kit com os campos preenchidos. A Segurança valida na tela Segurança do Trabalho.</div>' : ''}</div></div><div class="acoes">${souSeguranca() && !a.declinou_em ? `<button class="bt peq" data-act="osAbrir" data-id="${a.id}">Editar OS</button>` : ''}</div></div>` : ''}
    ${exigeLg || a.lg_movimentado_em ? `<div class="doc"><div><div class="dn">Movimentado no LG (recrutadora)</div><div class="dm">${chipLg}${a.lg_movimentado_por ? ` <span class="apoio" style="font-size:12px">por ${esc(a.lg_movimentado_por)}</span>` : ''}${!a.lg_movimentado_em && a.lg_avisado_em ? ` <span class="apoio" style="font-size:12px">aviso à analista em ${quando(a.lg_avisado_em)}</span>` : ''}${!a.lg_movimentado_em && exigeLg ? '<div class="motivo">Sem a movimentação para a etapa "Admissão" no LG, o DP não registra.</div>' : ''}</div></div><div class="acoes">${podeRs && !a.declinou_em && !a.matricula ? (a.lg_movimentado_em ? `<button class="bt peq fan" data-act="lgDesfazer" data-id="${a.id}">Desfazer</button>` : `<button class="bt peq pri" data-act="lgMovimentado" data-id="${a.id}">Movimentei no LG</button>`) : ''}</div></div>` : ''}
    <div class="doc"><div><div class="dn">Remuneração · posição ${esc(a.posicao || '—')}</div><div class="dm">${chipRemun(a)}</div></div><div class="acoes">${podeRem && !a.declinou_em && a.remuneracao !== 'ok' ? `<button class="bt peq" data-act="remun" data-id="${a.id}" data-v="ok">Validada</button><button class="bt peq" data-act="remun" data-id="${a.id}" data-v="ocupada">Posição ocupada</button>` : ''}</div></div>
    <div class="doc"><div><div class="dn">${a.tipo_contrato === 'Estágio' ? 'Cadastro do estágio e eSocial (S-2300)' : 'Matrícula e eSocial'}</div><div class="dm">${a.matricula ? chip('c-esc', '', a.matricula) + (a.esocial_enviado_em ? ' ' + chip('c-ok', 'ok', 'eSocial ' + dm(a.esocial_enviado_em) + (a.esocial_evento ? ' · ' + a.esocial_evento.split(' ')[0] : '')) : ' ' + chip('c-ag', 'ag', 'eSocial não registrado')) + (a.chamado_ti ? ' ' + chip('c-ag', '', 'Chamado TI ' + a.chamado_ti) : '') + ` <span class="apoio" style="font-size:12px">${esc(avisadas('matricula', a) || 'Nenhuma área com e-mail cadastrado para este momento.')}</span>` : a.etapa === 'registro' ? 'Pronto para registrar no LG' : 'Libera com documentos, ASO apto, remuneração ok' + (exigeLg ? ' e LG movimentado' : '')}</div></div><div class="acoes">${dp && a.etapa === 'registro' ? `<input id="matIn" class="sel-doc" style="width:110px" inputmode="numeric" placeholder="Matrícula" aria-label="Matrícula"><button class="bt peq pri" data-act="matricular" data-id="${a.id}">Registrar</button>` : ''}${dp && a.matricula && !a.declinou_em ? `<input type="date" id="esocialData" class="sel-doc" value="${a.esocial_enviado_em || ''}" aria-label="eSocial enviado em"><select id="esocialEv" class="sel-doc" aria-label="Evento"><option value="">Evento</option>${ESOCIAL.map((x) => `<option ${x === a.esocial_evento ? 'selected' : ''}>${x}</option>`).join('')}</select>${inputFicha('chamadoIn', 'Nº do chamado TI', `value="${esc(a.chamado_ti || '')}" style="width:130px"`)}<button class="bt peq" data-act="salvarRegistro" data-id="${a.id}">Salvar</button>${(a.ti_itens || []).length ? `<button class="bt peq" data-act="copiarChamado" data-id="${a.id}" title="Texto pronto para abrir o chamado no Jira">Copiar texto do chamado</button>` : ''}` : ''}</div></div>
    <div class="doc"><div><div class="dn">Para as áreas</div><div class="dm">Camisa ${esc(a.camisa || '—')} · ${a.fretado ? 'Fretado ' + esc(a.fretado) : 'Sem fretado'} · ${(a.ti_itens || []).length ? 'TI: ' + esc(a.ti_itens.join(', ')) + (a.cd_entrega ? ' (entrega no ' + esc(a.cd_entrega) + ')' : '') + (prazoChamadoTi(a) && !a.chamado_ti && !a.declinou_em ? ' · ' + chip(chamadoTiAtrasado(a) ? 'c-cr' : 'c-at', chamadoTiAtrasado(a) ? 'cr' : 'at', 'chamado TI até ' + dm(prazoChamadoTi(a))) : '') : 'Sem equipamento'}${a.usa_frota ? ' · Frota' : ''}${a.turno ? ' · ' + esc(a.turno) : ''}${a.cc_codigo || a.cc_nome ? ' · CC ' + esc([a.cc_codigo, a.cc_nome].filter(Boolean).join(' ')) : ''}${a.substituicao ? ' · Substitui ' + esc(a.pessoa_substituida || '(não informado)') : ''}</div></div></div>
   </div></div>
   ${ligado('dependentes_ativo') && tem('rs', 'admissao', 'saude', 'remuneracao') ? `<div class="caixa pilha"><div class="linha-flex"><h3 style="flex:1">Dependentes · ${plural(E.depsFicha.length, 'cadastrado', 'cadastrados')}</h3>${podeRs && !a.declinou_em ? `<button class="bt peq pri" data-act="fDepSalvar" data-id="${a.id}">Salvar dependentes</button>` : ''}</div>${blocoDeps(E.depsFicha, 'f', podeRs && !a.declinou_em)}<p class="apoio" style="font-size:12px">Os documentos de cada dependente entram na lista abaixo. Benefícios recebe o resumo no cadastro, na matrícula e quando o kit é assinado.</p></div>` : ''}
   ${docs.length ? `<div class="caixa pilha"><div class="linha-flex"><h3 style="flex:1">Documentos · ${r.ok} de ${r.total} conferidos</h3>${dp && !a.declinou_em && (r.pendentes || []).length ? `<button class="bt peq" data-act="cobrar" data-id="${a.id}">Cobrar ${plural(r.pendentes.length, 'pendência', 'pendências')}</button>` : ''}</div>
     <div class="docs">${docs.map((d) => `<div class="doc"><div><div class="dn">${esc(d.nome)}</div><div class="dm">${esc(d.dica || '')}</div>${d.motivo ? `<div class="motivo">Pedido porque: ${esc(d.motivo)}</div>` : ''}${d.obs ? `<div class="motivo">${esc(d.obs)}</div>` : ''}${d.atualizado_por && d.status !== 'falta' ? `<div class="motivo">${esc(d.atualizado_por)} · ${quando(d.atualizado_em)}</div>` : ''}${ehLaudo(d) ? (tem('admissao', 'saude') ? listaArquivos(a, d.tipo, arqs, { enviar: tem('admissao'), remover: false }) : '') + (dp && !tem('saude') ? '<div class="motivo">Só a Saúde abre o laudo e decide a situação.</div>' : '') : dp ? listaArquivos(a, d.tipo, arqs) : ''}</div><div class="acoes">${(ehLaudo(d) ? tem('saude') : dp) && !a.declinou_em ? selDoc(d) : stDoc[d.status]}</div></div>`).join('')}</div>
     ${dp ? '<p class="apoio" style="font-size:12px">Envie o arquivo na caixa do documento: a IA confere se é o documento certo, se está legível, se o nome bate com o cadastro e se falta frente ou verso. O mesmo arquivo não entra duas vezes. A situação continua editável: a analista sempre decide.</p>' : ''}</div>` : ''}
   ${dp && a.matricula ? `<div class="caixa pilha"><h3>Kit, benefícios e dossiê</h3>
     ${blocosKit(a).length ? `<div class="pilha" style="gap:4px"><span class="apoio" style="font-size:12.5px">Kit em blocos por área, cada um com o seu signatário pela Cimed (status gravado pela Esteira):</span><div class="linha-flex" style="flex-wrap:wrap">${blocosKit(a).map((b) => b.assinavel === false ? chip('c-ag', '', `${esc(b.titulo)} · só o colaborador assina${b.colab_em ? ' · ' + dm(new Date(b.colab_em)) : ''}`) : b.assinado_empresa_em ? chip('c-ok', 'ok', `${esc(b.titulo)} · Cimed: ${esc(b.empresa_nome || '')} ${dm(new Date(b.assinado_empresa_em))}${b.colab_em ? ' · colaborador ' + dm(new Date(b.colab_em)) : ''}`) : chip('c-at', 'at', `${esc(b.titulo)} · falta a Cimed${semSignatario(b.area) ? ' (sem signatário cadastrado)' : ''}${b.colab_em ? ' · colaborador ' + dm(new Date(b.colab_em)) : ''}`)).join('')}</div></div>` : a.kit_status !== 'nao' ? '<span class="apoio" style="font-size:12.5px">Kit sem status por bloco ainda: ele aparece quando a Cimed assina o primeiro bloco na Esteira.</span>' : ''}
     <div class="linha-flex">${E.termoFicha ? chip('c-ok', 'ok', 'Termo de Transparência assinado ' + dm(new Date(E.termoFicha.em))) + '<span class="apoio" style="font-size:12px">As respostas só o Compliance abre.</span>' : chip('c-ag', 'ag', 'Termo de Transparência ainda não assinado')}</div>
     ${E.opcFicha ? `<div class="pilha" style="gap:4px"><span class="apoio" style="font-size:12.5px">Escolhas de benefícios feitas no kit (${quando(E.opcFicha.recebido_em)}${E.opcFicha.avisado_em ? ', Benefícios avisado' : ''}):</span><div class="grade g2" style="gap:4px 16px;font-size:13px">${resumoOpcoes(E.opcFicha.dados) || '<span class="apoio">nenhuma escolha registrada</span>'}</div></div>` : '<span class="apoio" style="font-size:12.5px">As escolhas de benefícios chegam aqui quando o kit é assinado.</span>'}
     <div class="linha-flex" style="align-items:flex-start"><div style="flex:1;min-width:240px"><b>Dossiê</b><div class="apoio" style="font-size:12.5px">Baixa um .zip com o kit assinado e os documentos renomeados (MATRÍCULA - NOME/NN Documento), a ficha e as escolhas de benefícios. O laudo PCD fica fora. ${ligado('expurgo_exige_dossie') ? 'O descarte de 30 dias só roda depois disto.' : ''}</div>${a.arquivado_em ? `<div class="apoio" style="font-size:12.5px">Arquivado em ${quando(a.arquivado_em)} por ${esc(a.arquivado_por || '')}.</div>` : ''}</div><button class="bt ${a.arquivado_em ? '' : 'pri'}" data-act="dossie" data-id="${a.id}">${a.arquivado_em ? 'Exportar de novo' : 'Exportar dossiê (.zip)'}</button></div></div>` : ''}
   ${dp && (E.acessosFicha || []).length ? `<div class="caixa pilha" style="gap:6px"><h3>Quem abriu documentos</h3>${E.acessosFicha.map((x) => `<span class="apoio" style="font-size:12.5px">${esc((docs.find((d) => d.tipo === x.tipo) || {}).nome || x.tipo)} · ${esc(x.quem)} · ${quando(x.em)}</span>`).join('')}</div>` : ''}
   <div class="caixa pilha"><h3>Linha do tempo</h3><div class="tempo">${evs.map((h) => `<div class="ev"><span class="pt"></span><div>${esc(textoEvento(h.texto))}<small>${quando(h.em)} · ${esc(h.quem)}</small></div></div>`).join('') || '<span class="apoio">Sem registros.</span>'}</div>
    ${tem('rs', 'admissao', 'saude', 'remuneracao') ? `<form id="fObs" class="linha-flex"><input id="obsTx" maxlength="1000" placeholder="Escrever uma observação (fica com seu nome e a hora)" style="flex:1;min-width:200px;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup)"><button class="bt" type="submit">Anotar</button></form>` : ''}</div>
   <div class="caixa pilha"><h3>Dados da vaga</h3><div class="grade g2" style="gap:6px 16px;font-size:13px">${[['Escala', a.escala], ['Turno', a.turno], ['Posição', a.posicao], ['Processo seletivo', a.processo], ['Centro de custo', [a.cc_codigo, a.cc_nome].filter(Boolean).join(' · ')], ['Vaga', a.substituicao ? 'Substituição de ' + (a.pessoa_substituida || '?') : 'Aumento de quadro'], ['Requisitante', a.requisitante], ['E-mail do requisitante', a.requisitante_email], ['Analista R&S', a.analista_rs], ['Analista DP', a.analista_dp], ['CBO', a.cbo], ['Estado civil', a.estado_civil], ['Filhos', a.filhos], ['Formato da integração', a.formato_integ === 'online' ? 'On-line' : 'Presencial'], ['Chamado de TI', a.chamado_ti]].map(([k, v]) => `<div><span class="apoio">${esc(k)}:</span> <b>${esc(v || '—')}</b></div>`).join('')}</div>
    ${podeRs && !a.declinou_em && !a.matricula ? `<form id="fDecl" class="linha-flex"><input id="declTx" placeholder="Motivo do declínio" style="flex:1;min-width:200px;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup)"><label class="apoio" style="font-size:12.5px;display:flex;align-items:center;gap:6px"><input type="checkbox" id="declSens"> Motivo sensível: as áreas recebem só "cancelada"</label><button class="bt" type="submit">Registrar declínio</button></form>` : ''}
    ${tem('admissao') && !a.declinou_em && a.matricula ? `<form id="fNaoVeio" class="linha-flex"><input id="naoVeioTx" placeholder="Não compareceu ou desistiu depois da matrícula: o que aconteceu?" style="flex:1;min-width:200px;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup)"><button class="bt" type="submit">Registrar não comparecimento</button></form>` : ''}</div>
  </aside></div>`;

  document.querySelectorAll('.gaveta .sel-doc[data-doc]').forEach((s) => s.addEventListener('change', async (e) => {
    const st = e.target.value; let obs = null;
    if (st === 'reenviar') { obs = await perguntar('Por que precisa reenviar?', 'Ex.: foto desfocada, falta o verso, comprovante antigo'); if (obs === null) { e.target.value = docs.find((d) => d.tipo === e.target.dataset.doc).status; return; } }
    const r2 = await tentar(() => api.marcarDoc(a.id, e.target.dataset.doc, st, obs));
    if (r2 && !r2.length) toast('Você não tem permissão para mudar este documento.', true);
    await recarregar(); abrirFicha(a.id);
  }));
  $('#fDeps')?.addEventListener('input', (e) => { const el = e.target; if (el.dataset.dep == null) return; const d = E.depsFicha[+el.dataset.dep]; if (d) d[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value; });
  $('#fObs')?.addEventListener('submit', async (e) => { e.preventDefault(); const t = $('#obsTx').value.trim(); if (!t) return; await tentar(() => api.anotar(a.id, t), 'Anotado'); abrirFicha(a.id); });
  $('#fNaoVeio')?.addEventListener('submit', async (e) => {
    e.preventDefault(); const t = $('#naoVeioTx').value.trim(); if (!t) { toast('Escreva o que aconteceu.', true); return; }
    if (await tentar(() => api.atualizar(a.id, { declinou_motivo: 'Não compareceu depois da matrícula: ' + t, declinou_em: new Date().toISOString() }), 'Não comparecimento registrado')) {
      await tentar(() => api.anotar(a.id, 'Pendência no LG: matrícula ' + a.matricula + ' já foi gerada. ' + (a.esocial_enviado_em
        ? 'eSocial enviado em ' + dm(a.esocial_enviado_em) + ': se não houve prestação de serviço, enviar a exclusão (S-3000); se só mudou a data, retificar o S-2200.'
        : 'Conferir se o eSocial foi enviado: se foi e não houve prestação de serviço, enviar a exclusão (S-3000); se só mudou a data, retificar o S-2200.')));
      await recarregar(); abrirFicha(a.id);
    }
  });
  $('#fDecl')?.addEventListener('submit', async (e) => {
    e.preventDefault(); const t = $('#declTx').value.trim(); if (!t) { toast('Escreva o motivo do declínio.', true); return; }
    if (await tentar(() => api.atualizar(a.id, { declinou_motivo: t, declinou_em: new Date().toISOString(), declinou_sensivel: !!$('#declSens')?.checked }), 'Declínio registrado')) { await recarregar(); abrirFicha(a.id); }
  });
}

// Dossiê: baixa kit e documentos, renomeia e empacota num .zip; depois marca arquivado
async function exportarDossie(id, bt) {
  const a = adm(id); if (!a) return;
  bt.disabled = true; const rot = bt.textContent; bt.textContent = 'Montando o dossiê…';
  try {
    const d = await api.dossie(id);
    const pasta = nomeArquivoSeguro((a.matricula ? a.matricula + ' - ' : '') + a.nome) + '/';
    const entradas = []; const enc = new TextEncoder(); let n = 0, kit = false; const falhas = [];
    const ext = (m, nome) => ({ 'application/pdf': 'pdf', 'image/png': 'png', 'image/webp': 'webp', 'image/jpeg': 'jpg' }[m] || (String(nome || '').split('.').pop() || 'bin'));
    if (d.kit && d.kit.pdf) {
      try { const b = await api.baixarKitArquivo(d.kit.pdf); entradas.push({ nome: pasta + nomeArquivoSeguro('KIT ADMISSAO - ' + a.nome + (a.matricula ? ' - ' + a.matricula : '') + (d.kit.colab_assinou && d.kit.empresa_assinou ? ' - ASSINADO' : '')) + '.pdf', bytes: new Uint8Array(await b.arrayBuffer()) }); kit = true; }
      catch (e) { falhas.push('kit: ' + e.message); }
    }
    let i = 0;
    for (const f of d.arquivos || []) {
      i++; bt.textContent = `Baixando documento ${i} de ${(d.arquivos || []).length}…`;
      try { const b = await api.baixarArquivo(f.id, f.caminho); entradas.push({ nome: pasta + String(i).padStart(2, '0') + ' ' + nomeArquivoSeguro(f.documento + (f.lado && f.lado !== 'unico' ? ' ' + f.lado : '')) + '.' + ext(f.mime, f.nome), bytes: new Uint8Array(await b.arrayBuffer()) }); n++; }
      catch (e) { falhas.push((f.documento || f.tipo) + ': ' + e.message); }
    }
    const ficha = Object.assign({}, d.ficha, { dependentes: d.dependentes, opcoes_beneficios: d.opcoes_beneficios, termo_transparencia: d.termo_transparencia, gerado_em: d.gerado_em, por: d.por, falhas });
    entradas.push({ nome: pasta + '00 Ficha.json', bytes: enc.encode(JSON.stringify(ficha, null, 2)) });
    entradas.push({ nome: pasta + '00 Linha do tempo.txt', bytes: enc.encode((d.eventos || []).map((e) => quando(e.em) + ' · ' + e.quem + ' · ' + e.texto).join('\r\n')) });
    baixarBlob(nomeArquivoSeguro('DOSSIE - ' + a.nome + (a.matricula ? ' - ' + a.matricula : '')) + '.zip', zipar(entradas));
    await api.dossieArquivado(id, n, kit, { arquivos: n, kit, falhas, entradas: entradas.map((e) => e.nome) });
    toast(`Dossiê exportado: ${plural(n, 'documento', 'documentos')}${kit ? ' e o kit' : ', sem kit assinado'}.` + (falhas.length ? ' Não baixou: ' + falhas.join('; ') : ''), falhas.length > 0);
    await recarregar(); abrirFicha(id);
  } catch (e) { toast('Não consegui montar o dossiê: ' + e.message, true); bt.disabled = false; bt.textContent = rot; }
}

// "Etapa: pendencia → docs_ok" vira os nomes que aparecem no painel
function textoEvento(t) {
  const nome = (id) => (id === 'declinou' ? 'Declinou' : ETAPA[id]?.n || id);
  return String(t).replace(/^Etapa: (\w+) → (\w+)$/, (_, a, b) => `Etapa: ${nome(a)} → ${nome(b)}`)
    .replace(/^ASO: (apto_restricao|aguardando|agendado|apto|inapto)/, (_, s) => 'ASO: ' + { aguardando: 'a agendar', agendado: 'agendado', apto: 'apto', apto_restricao: 'apto com restrição', inapto: 'inapto' }[s])
    .replace(/^Remuneração: ok$/, 'Remuneração validada').replace(/^Remuneração: ocupada$/, 'Remuneração: posição ocupada');
}

// pergunta dentro da página (o navegador do Claude e alguns ambientes bloqueiam prompt())
function perguntar(tit, dica) {
  return new Promise((ok) => {
    const div = document.createElement('div'); div.className = 'modal-veu';
    div.innerHTML = `<form class="modal" role="dialog" aria-modal="true"><h2>${esc(tit)}</h2><input id="pTx" class="sel-doc" style="min-height:42px" placeholder="${esc(dica)}"><div class="linha-flex" style="justify-content:flex-end"><button type="button" class="bt" id="pNao">Cancelar</button><button class="bt pri">Confirmar</button></div></form>`;
    document.body.appendChild(div); div.querySelector('#pTx').focus();
    div.querySelector('#pNao').onclick = () => { div.remove(); ok(null); };
    div.querySelector('form').onsubmit = (e) => { e.preventDefault(); const v = div.querySelector('#pTx').value.trim(); div.remove(); ok(v || 'A analista pediu um novo envio'); };
  });
}

// ----------------------------------------------------------- documentos
VIEWS.documentos = () => {
  const fila = vivas().filter((a) => ['aguardando', 'pendencia', 'conferencia'].includes(a.etapa))
    .sort((x, y) => (resumo(y).analise - resumo(x).analise) || x.data_inicio.localeCompare(y.data_inicio));
  return `<div class="cab"><div class="t"><span class="olho">ADM de Pessoal</span><h1>Documentos</h1><p class="apoio">Quem ainda não está com a documentação completa, do corte mais próximo para o mais distante. Abra a ficha para enviar os arquivos: a IA confere se cada um é o documento certo e a analista confirma. A cobrança automática sai todo dia útil às 9h, uma vez por pessoa, só com o que falta e até o corte.</p></div></div>
  ${fila.length ? `<div class="tab-wrap"><table class="t-docs"><thead><tr><th>Pessoa</th><th>Turma</th><th>Etapa</th><th>Conferidos</th><th>O que falta</th><th>Cobranças</th></tr></thead><tbody>
  ${fila.map((a) => { const r = resumo(a); return `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.local)}</span></td><td class="num">${dmd(a.data_inicio)}${foraCorte(a) ? '<br>' + chip('c-cr', 'cr', 'Fora do corte') : ''}</td><td>${chipEtapa(a)}</td><td class="num">${r.ok}/${r.total}${r.analise ? '<br>' + chip('c-info', 'info', r.analise + ' para conferir') : ''}</td><td style="font-size:12.5px">${esc((r.pendentes || []).join(', ') || '—')}</td><td><div class="cobr">${(r.pendentes || []).length && a.email ? `<button class="bt peq" data-act="cobrar" data-id="${a.id}">Cobrar</button>` : ''}<span class="apoio">${a.cobrancas ? plural(a.cobrancas, 'enviada', 'enviadas') + (a.ultima_cobranca ? ' · ' + quando(a.ultima_cobranca) : '') : 'nenhuma ainda'}</span></div></td></tr>`; }).join('')}
  </tbody></table></div>` : '<div class="vazio">Ninguém com documentação pendente.</div>'}`;
};

// --------------------------------------------------------- remuneração
VIEWS.remuneracao = () => {
  const l = vivas().filter((a) => a.remuneracao !== 'ok' && !a.matricula).sort((x, y) => x.data_inicio.localeCompare(y.data_inicio));
  return `<div class="cab"><div class="t"><span class="olho">Remuneração</span><h1>Validar remuneração</h1><p class="apoio">Posições esperando validação, da turma mais próxima para a mais distante. A matrícula só libera depois desta validação. O combinado é responder em até 1 dia útil.</p></div></div>
  ${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Cargo</th><th>Posição</th><th>Turma</th><th>Cadastro</th><th></th></tr></thead><tbody>${l.map((a) => `<tr><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.local)}</span></td><td>${esc(a.cargo)}</td><td class="num">${esc(a.posicao || '—')}</td><td class="num">${dmd(a.data_inicio)}</td><td class="num">${quando(a.criado_em)}</td><td><div class="linha-flex">${a.remuneracao === 'ocupada' ? chip('c-cr', 'cr', 'Ocupada') : ''}<button class="bt peq pri" data-act="remun" data-id="${a.id}" data-v="ok">Validada</button>${a.remuneracao !== 'ocupada' ? `<button class="bt peq" data-act="remun" data-id="${a.id}" data-v="ocupada">Posição ocupada</button>` : ''}</div></td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nada esperando validação.</div>'}`;
};

// ---------------------------------------------------------------- saúde
function laudosParaSaude() {
  const l = E.admissoes.filter((a) => a.pcd && a.pcd !== 'Não' && !a.declinou_em && !a.expurgado_em);
  if (!l.length) return '<div class="vazio">Nenhum candidato PCD em andamento.</div>';
  return `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Deficiência</th><th>Integração</th><th></th></tr></thead><tbody>${l.map((a) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.cargo)} · ${esc(a.local)}</span></td><td>${esc(a.pcd)}</td><td class="num">${dmd(a.data_inicio)}</td><td><button class="bt peq" data-act="ficha" data-id="${a.id}">Abrir laudo</button></td></tr>`).join('')}</tbody></table></div>`;
}
function filaSaude() {
  const l = vivas().filter((a) => !a.matricula && a.kit_status !== 'concluido').sort((x, y) => (x.data_inicio + (asoOk(x) ? '1' : '0') + x.local).localeCompare(y.data_inicio + (asoOk(y) ? '1' : '0') + y.local));
  if (!l.length) return '<div class="vazio">Ninguém esperando exame.</div>';
  const corte = (a) => a.corte_documentos ? Math.ceil((new Date(a.corte_documentos).getTime() - Date.now()) / 86400000) : null;
  return `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Cargo · local</th><th>Integração</th><th>Contato</th><th>Exame</th><th>Dias até o corte</th><th>LG</th></tr></thead><tbody>
  ${l.map((a) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b>${a.pcd !== 'Não' ? ' ' + chip('c-info', 'info', 'PCD') : ''}<br><span class="apoio num" style="font-size:12px">${a.cpf ? 'CPF ' + fmtCpf(a.cpf) : ''}</span></td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.local)}${a.setor ? ' · ' + esc(a.setor) : ''}</span></td><td class="num">${dmd(a.data_inicio)}</td><td>${esc(a.telefone || '—')}<br><span class="apoio" style="font-size:12px">${esc(a.email || '')}</span></td><td>${chipAso(a)}${a.aso_restricao ? '<br><span class="apoio" style="font-size:12px">' + esc(a.aso_restricao) + '</span>' : ''}</td><td class="num">${corte(a) == null ? '—' : corte(a) < 0 ? chip('c-cr', 'cr', 'corte passou') : corte(a) <= 2 && !asoOk(a) ? chip('c-at', 'at', corte(a) + ' d') : corte(a) + ' d'}</td><td>${asoOk(a) ? (a.lg_movimentado_em ? chip('c-ok', 'ok', 'LG ' + dm(a.lg_movimentado_em)) : chip('c-ag', 'ag', 'LG pendente')) : '—'}</td></tr>`).join('')}
  </tbody></table></div>`;
}
VIEWS.saude = () => `<div class="cab"><div class="t"><span class="olho">Saúde ocupacional</span><h1>Exame admissional</h1><p class="apoio">Quem precisa de exame, por turma. Clique na pessoa para agendar (data, hora e local: o candidato recebe o e-mail), registrar apto, apto com restrição ou inapto. Apto libera a matrícula e pede à recrutadora a movimentação no LG; inapto avisa o R&S e a Admissão, e a Saúde decide na ficha se remarca o exame ou cancela a admissão. O ASO oficial continua no SOC.</p></div></div>
 <div class="caixa pilha"><h2>Fila de agendamento</h2>${filaSaude()}</div>
 <div class="caixa pilha"><h2>Laudos PCD</h2><p class="apoio">Só a Saúde abre o laudo médico e decide a situação dele. A Admissão vê apenas se foi entregue.</p>${laudosParaSaude()}</div>`;
// ----------------------------------------------------------- Segurança do Trabalho
// Ordem de Serviço por pessoa (reunião de 15/10/2026): a matriz cargo → riscos/EPI/EPC
// pré-preenche, a Segurança valida, e a Esteira gera a OS no bloco Segurança do kit.
const CAMPOS_OS = [['os_atividades', 'Descrição das atividades'], ['os_riscos', 'Riscos da função (PGR)'], ['os_epi', 'EPI'], ['os_epc', 'EPC e medidas de controle'], ['os_nrs', 'Treinamentos e NRs (NR-6, NR-12, NR-35, NR-20…)']];
function abrirOs(id) {
  const a = adm(id); if (!a) return;
  const pode = souSeguranca() && !a.declinou_em;
  $('#camada').innerHTML = `<div class="veu" data-act="fechar"><aside class="gaveta" data-pare="1"><div class="pilha">
   <div class="linha-flex" style="align-items:flex-start"><div class="pilha" style="flex:1;gap:6px"><span class="olho">Ordem de Serviço · Segurança do Trabalho</span><h2 style="font-size:22px">${esc(titulo(a.nome))}</h2><span class="apoio">${esc(a.cargo)}${a.setor ? ' · ' + esc(a.setor) : ''} · ${esc(a.local)} · integração ${dmd(a.data_inicio)}${a.matricula ? ' · matrícula ' + esc(a.matricula) : ''}</span><div class="linha-flex">${chipOs(a)}${a.os_matriz_versao ? chip('c-ag', '', 'matriz v' + a.os_matriz_versao) : ''}${a.nr12 ? chip('c-ag', '', 'NR-12') : ''}</div></div><button class="bt" data-act="fechar">Fechar</button></div>
   <div class="caixa pilha"><div class="campo"><label for="os_cbo">CBO</label><input id="os_cbo" value="${esc(a.cbo || '')}" ${pode ? '' : 'readonly'} style="max-width:160px"></div>
   ${CAMPOS_OS.map(([k, l]) => `<div class="campo"><label for="os_${k}">${esc(l)}</label><textarea id="os_${k}" rows="3" ${pode ? '' : 'readonly'} style="width:100%;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup);font:inherit">${esc(a[k] || '')}</textarea></div>`).join('')}
   <p class="apoio" style="font-size:12.5px">A parte fixa da OS (objetivo, obrigações do empregado pela NR-1, CLT cap. V, orientações gerais) já vem no modelo da Esteira. Aqui vai só o que muda por função. O colaborador assina a OS no kit; ${a.os_status === 'validada' ? 'esta OS está validada' : 'sem validação a Esteira avisa e não gera a OS no kit'}.</p>
   ${pode ? `<div class="linha-flex" style="justify-content:flex-end"><button class="bt" data-act="osSugerir" data-id="${a.id}" title="Preenche os campos vazios pela matriz cargo → riscos">Preencher pela matriz</button><button class="bt" data-act="osSalvar" data-id="${a.id}">Salvar</button><button class="bt pri" data-act="osValidar" data-id="${a.id}">Salvar e validar</button></div>` : ''}</div>
  </div></aside></div>`;
}
VIEWS.seguranca = () => {
  const l = vivas().filter((a) => temOs(a)).sort((x, y) => (x.data_inicio + (x.os_status === 'validada' ? '1' : '0') + x.local).localeCompare(y.data_inicio + (y.os_status === 'validada' ? '1' : '0') + y.local));
  const unidades = E.locais.map((x) => x.nome);
  const optU = (sel) => unidades.map((u) => `<option ${u === sel ? 'selected' : ''}>${esc(u)}</option>`).join('');
  const linha = (m, i) => `<div class="area" data-mos="${i}"><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(m.unidade)}</select></div><div class="ac"><label class="olho">Setor (opcional)</label><input class="sel-doc" data-f="setor" value="${esc(m.setor || '')}"></div><div class="ac"><label class="olho">Cargo</label><input class="sel-doc" data-f="cargo" value="${esc(m.cargo || '')}" placeholder="como no cadastro"></div><div class="ac"><label class="olho">GHE</label><input class="sel-doc" data-f="ghe" value="${esc(m.ghe || '')}" style="width:90px"></div>
    ${[['atividades', 'Atividades'], ['riscos', 'Riscos'], ['epi', 'EPI'], ['epc', 'EPC'], ['nrs', 'NRs']].map(([k, t]) => `<div class="ac" style="flex-basis:100%"><label class="olho">${t}</label><textarea class="sel-doc" data-f="${k}" rows="2" style="width:100%;font:inherit">${esc(m[k] || '')}</textarea></div>`).join('')}
    <div class="ac"><span class="apoio" style="font-size:12px">${m.id ? 'v' + (m.versao || 1) + (m.atualizado_por ? ' · ' + esc(m.atualizado_por) : '') : 'nova linha'}</span></div><div class="ac"><button class="bt peq pri" data-act="mosSalvar" data-i="${i}">Salvar</button>${m.id ? `<button class="bt peq fan" data-act="mosApagar" data-id="${m.id}">Remover</button>` : ''}</div></div>`;
  return `<div class="cab"><div class="t"><span class="olho">Segurança do Trabalho</span><h1>Ordem de Serviço</h1><p class="apoio">Uma OS por pessoa, em todas as unidades do Brasil (decisão de 20/10/2026). A matriz abaixo (cargo → atividades, riscos, EPI, EPC, NRs, por unidade e setor) pré-preenche a OS no cadastro; a Segurança confere, ajusta o que for diferente e valida. A OS validada entra no kit admissional, no bloco Segurança, e o colaborador assina na integração. A Segurança recebe o aviso no cadastro (previsão) e na matrícula (confirmação).</p></div></div>
  <div class="caixa pilha"><h2>Quem precisa de OS</h2>${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Cargo · setor</th><th>Unidade</th><th>Integração</th><th>CBO</th><th>OS</th><th></th></tr></thead><tbody>
   ${l.map((a) => `<tr><td><b>${esc(titulo(a.nome))}</b>${a.nr12 ? ' ' + chip('c-ag', '', 'NR-12') : ''}</td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.setor || '')}</span></td><td>${esc(a.local)}</td><td class="num">${dmd(a.data_inicio)}</td><td class="num">${esc(a.cbo || '—')}</td><td>${chipOs(a)}</td><td><button class="bt peq ${a.os_status === 'validada' ? '' : 'pri'}" data-act="osAbrir" data-id="${a.id}">${a.os_status === 'validada' ? 'Ver OS' : 'Preencher e validar'}</button></td></tr>`).join('')}
   </tbody></table></div>` : '<div class="vazio">Ninguém de fábrica ou CD em admissão.</div>'}</div>
  <div class="caixa pilha"><h2>Matriz cargo → riscos, EPI e EPC</h2><p class="apoio">Compilada pela Segurança a partir do PGR de cada unidade, por grupo homogêneo de exposição. Linha com setor vale só para aquele setor; sem setor, vale para o cargo na unidade inteira. Cada alteração vira uma versão nova, e a OS impressa diz de qual versão saiu.</p>
   ${(E.matrizOs || []).map(linha).join('')}${linha({ unidade: unidades[0] || '' }, (E.matrizOs || []).length)}</div>`;
};

// ----------------------------------------------------------- assinatura
// Esteira de Admissão: a mesma ferramenta de sempre, dentro da plataforma.
// Sem publicar on-line, os kits ficam no navegador do aparelho em que foram gerados.
// Onde está cada kit da turma (Josianny, 20/10/2026): trilha gerado → publicado → Cimed assinou →
// colaborador assinou → concluído, e com quem está. Lê admissoes.kit_status/kit_blocos e kits_online.
function ondeEstaCadaKit(turma) {
  const lista = vivas().filter((a) => a.matricula && (turma === 'todas' || a.data_inicio === turma));
  if (!lista.length) return '';
  const online = Object.fromEntries((E.kitsOnline || []).filter((k) => k.admissao_id && !k.cancelado_em).map((k) => [k.admissao_id, k]));
  const sig = (area) => { try { const m = JSON.parse(E.params.signatarios_por_area || '{}'); return (m[area] || []).map((s) => String(s.nome || '').split(' ')[0]).filter(Boolean); } catch (e) { return []; } };
  const linha = (a) => {
    const ko = online[a.id]; const bl = blocosKit(a); const falta = faltaAssinar(a);
    const gerado = a.kit_status !== 'nao' || !!ko || bl.length > 0;
    const publicado = !!ko;
    const cimedOk = ['empresa', 'concluido'].includes(a.kit_status) || (bl.length > 0 && !falta.length && bl.some((b) => b.assinado_empresa_em)) || !!(ko && ko.empresa_assinou);
    const cimedParte = !cimedOk && bl.some((b) => b.assinado_empresa_em);
    const colab = a.kit_status === 'concluido' || !!(ko && ko.colab_assinou) || bl.some((b) => b.colab_em);
    const concluido = a.kit_status === 'concluido';
    // etapa feita: "Cimed assinou"; etapa atual (amarela): "Pendente: colaborador assinar"
    const passos = [['Gerado', gerado, false, 'gerar'], ['Publicado', publicado, false, 'publicar'], ['Cimed assinou', cimedOk, cimedParte, 'Cimed assinar'], ['Colaborador assinou', colab, false, 'colaborador assinar'], ['Concluído', concluido, false, 'concluir']];
    let ultimo = -1; passos.forEach((x, i) => { if (x[1]) ultimo = i; });
    const trilha = `<div class="trilha-kit">${passos.map((x, i) => { const atual = !x[1] && i === ultimo + 1; return `<span class="tp${x[1] ? ' ok' : x[2] ? ' parte' : atual ? ' atual' : ''}"><i>${x[1] ? '✓' : i + 1}</i>${x[1] ? x[0] : x[2] ? 'Cimed assinou em parte' : atual ? 'Pendente: ' + x[3] : x[0]}</span>`; }).join('')}</div>`;
    let com;
    if (concluido) com = 'Admissão · entregar a via';
    else if (!gerado) com = 'Admissão · gerar o kit na Esteira';
    else if (!cimedOk) {
      const areasPend = bl.filter((b) => b.assinavel !== false && !b.assinado_empresa_em).map((b) => b.area);
      const nomes = Array.from(new Set((areasPend.length ? areasPend : ['dp']).flatMap(sig)));
      com = (nomes.length ? nomes.join(', ') : 'Cimed') + ' · assinar ' + (falta.length ? falta.join(', ') : 'o kit');
    } else if (!colab) com = ko && ko.token_hash && new Date(ko.expira_em) > new Date() ? 'colaborador · assinar pelo celular (link ativo)' : 'Admissão · mandar o link ou entregar o tablet';
    else com = 'Admissão';
    return `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b><br><small class="apoio">${esc(a.cargo || '')}</small></td><td>${trilha}</td><td>${esc(com)}</td></tr>`;
  };
  return `<div class="caixa" style="margin-bottom:12px"><b>Onde está cada kit</b><p class="apoio" style="font-size:12.5px;margin:4px 0 8px">O kit vai para quem assina pela Cimed e volta para a Admissão sem ninguém clicar; esta lista mostra a etapa de cada um (${esc(turma === 'todas' ? 'todas as turmas' : dmd(turma))}).</p>
  <div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Etapa</th><th>Está com</th></tr></thead><tbody>${lista.map(linha).join('')}</tbody></table></div></div>`;
}
// quem só assina pela Cimed (papel signatario, sem admissao) abre a Esteira na tela própria de assinatura
const modoEsteira = () => (tem('signatario') && !tem('admissao') ? '&modo=signatario' : '');
VIEWS.kit = () => {
  const ts = turmas();
  const prox = E.turmaKit || ts.find((t) => deIso(t) >= deIso(iso(new Date()))) || ts[ts.length - 1] || '';
  const prontos = vivas().filter((a) => a.matricula && a.data_inicio === prox).length;
  if (modoEsteira()) return `<div class="cab"><div class="t"><span class="olho">Kit admissional</span><h1>Assinatura pela Cimed</h1><p class="apoio">Os kits publicados pela Admissão chegam aqui sozinhos. Leia o kit se quiser e assine com a sua assinatura; ela volta para a Admissão na hora.</p></div>
    <a class="bt" href="${(window.ESTEIRA_URL || 'esteira.html') + '?v=27' + modoEsteira()}" target="_blank" rel="noopener">Abrir em tela cheia</a></div>
  ${ondeEstaCadaKit('todas')}`;
  return `<div class="cab"><div class="t"><span class="olho">Kit admissional</span><h1>Esteira de assinatura</h1><p class="apoio">A mesma Esteira de Admissão de sempre: junta os PDFs do LG por pessoa, a Cimed assina primeiro e o colaborador assina na tela. Com <b>Publicar on-line</b>, o kit vai para o armazenamento privado da plataforma: a Cimed assina de qualquer aparelho e o colaborador assina pelo QR, no celular dele, depois de confirmar CPF e data de nascimento. Sem publicar, o kit fica só neste aparelho, como antes.</p></div>
    <a class="bt" href="${(window.ESTEIRA_URL || 'esteira.html') + '?v=27' + modoEsteira()}" target="_blank" rel="noopener">Abrir em tela cheia</a></div>
  ${faixaBlocos()}
  ${ondeEstaCadaKit(prox)}
  ${tem('admissao') && ts.length ? `<div class="faixa"><p><b>Lista de admitidos para a Esteira.</b> Copie os matriculados de uma turma e cole no passo 1 da Esteira, em "Colar da planilha". Nome, CPF, matrícula, unidade e cargo vão no formato que ela lê.</p>
    <label class="olho" for="turmaKit">Turma</label><select id="turmaKit" class="sel-doc">${ts.map((t) => `<option value="${t}" ${t === prox ? 'selected' : ''}>${dmd(t)}</option>`).join('')}</select>
    <button class="bt ama" data-act="copiarEsteira" ${prontos ? '' : 'disabled'}>Copiar ${plural(prontos, 'matriculado', 'matriculados')}</button></div>` : ''}`;
};
// O iframe nasce uma vez e fica: trocar de tela só esconde. Recarregar
// a Esteira perderia os PDFs que já foram jogados nela.
POS.kit = () => {
  $('#turmaKit')?.addEventListener('change', (e) => { E.turmaKit = e.target.value; render(); });
  const host = $('#esteiraHost');
  // o "ver como" troca entre a Esteira completa e a tela do signatário: o iframe nasce de novo
  const atual = host.firstChild;
  if (atual && String(atual.src || '').includes('modo=signatario') !== !!modoEsteira()) host.innerHTML = '';
  if (!host.firstChild) {
    host.innerHTML = `<iframe class="esteira" src="${(window.ESTEIRA_URL || 'esteira.html') + '?v=27' + modoEsteira()}" title="Esteira de Admissão"></iframe>`;
    // a Esteira lê os mesmos parâmetros do Supabase; mandar daqui cobre o teste sem banco e evita esperar
    const fr = host.firstChild;
    fr.addEventListener('load', () => {
      try { const alvo = new URL(fr.src, location.href).origin; fr.contentWindow.postMessage({ tipo: 'cimed-parametros', signatarios_cimed: E.params.signatarios_cimed, signatarios_por_area: E.params.signatarios_por_area, termo_transparencia_ativo: E.params.termo_transparencia_ativo, os_ativa: E.params.os_ativa, whatsapp_ativo: E.params.whatsapp_ativo }, alvo === 'null' ? '*' : alvo); } catch (e) { /* iframe sem acesso */ }
    });
  }
  host.hidden = false;
};

// Formato do "Colar da planilha" da Esteira: NOME CPF E-MAIL NASCIMENTO MATRICULA UNIDADE CARGO SETOR CC ADMISSAO
// (o centro de custo preenche a "área" e a data de admissão a "data" no cabeçalho do Termo de Transparência)
function listaParaEsteira(turma) {
  const unidade = (l) => ({ 'Fernão Dias': 'Pouso Alegre (Fernão Dias)', 'Gráfica SSBV': 'SSBV', 'Angélica': 'Escritório SP', 'Faria Lima': 'Escritório SP' }[l] || l);
  const fmtN = (d) => (d ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : '');
  const linhas = vivas().filter((a) => a.matricula && a.data_inicio === turma)
    .map((a) => [a.nome, fmtCpf(a.cpf), a.email || '', fmtN(a.nasc), a.matricula, unidade(a.local), a.cargo, a.setor || '', [a.cc_codigo, a.cc_nome].filter(Boolean).join(' - '), fmtN(a.data_inicio),
      a.cbo || '', temOs(a) ? a.os_status : 'nao_aplica', a.os_atividades || '', a.os_riscos || '', a.os_epi || '', a.os_epc || '', a.os_nrs || '', a.os_validada_por || '', a.os_matriz_versao || '', a.telefone || ''].map((x) => String(x).replace(/[\t\r\n]+/g, ' ')).join('\t'));
  return ['NOME\tCPF\tE-MAIL\tNASCIMENTO\tMATRICULA\tUNIDADE\tCARGO\tSETOR\tCC\tADMISSAO\tCBO\tOS_STATUS\tOS_ATIVIDADES\tOS_RISCOS\tOS_EPI\tOS_EPC\tOS_NRS\tOS_VALIDADA_POR\tOS_MATRIZ_VERSAO\tWHATSAPP', ...linhas].join('\n');
}

// ------------------------------------------------------------- entregas
// ----------------------------------------------------------- TI
// Previsibilidade desde o cadastro (decisão de 20/10/2026): a TI vê quem vai entrar e com qual
// equipamento antes da matrícula, registra o chamado e vê os cancelamentos para desfazer.
VIEWS.ti = () => {
  const comTi = (a) => (a.ti_itens || []).length > 0;
  const hoje = new Date(); const limite = new Date(hoje.getTime() - 30 * 86400000);
  const l = vivas().filter((a) => comTi(a) && deIso(a.data_inicio) >= limite).sort((x, y) => (x.data_inicio + x.local + x.nome).localeCompare(y.data_inicio + y.local + y.nome));
  const canc = E.admissoes.filter((a) => a.declinou_em && comTi(a) && (a.chamado_ti || a.matricula) && new Date(a.declinou_em) >= limite).sort((x, y) => y.declinou_em.localeCompare(x.declinou_em));
  const em7 = l.filter((a) => deIso(a.data_inicio) - hoje < 7 * 86400000 && deIso(a.data_inicio) >= hoje).length;
  const semChamado = l.filter((a) => a.matricula && !a.chamado_ti).length;
  const atras = l.filter(chamadoTiAtrasado).length;
  const situacao = (a) => a.kit_status === 'concluido' ? chip('c-ok', 'ok', 'Kit assinado') : a.matricula ? chip('c-ag', 'ag', 'Pedido · matrícula ' + a.matricula) : chip('c-info', 'info', 'Previsão (' + ((ETAPAS.find((e) => e.id === a.etapa) || {}).n || a.etapa) + ')');
  const linha = (a) => `<tr><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.cargo)}${a.setor ? ' · ' + esc(a.setor) : ''} · ${esc(a.local)}</span></td><td class="num">${dmd(a.data_inicio)} ${esc(a.hora_integ || '')}<br><span class="apoio" style="font-size:12px">${esc(a.integ || '')}${a.formato_integ === 'online' ? ' · on-line' : ''}</span></td><td style="font-size:12.5px">${esc((a.ti_itens || []).join(', '))}${a.cd_entrega ? `<br><span class="apoio" style="font-size:12px">entrega no ${esc(a.cd_entrega)}</span>` : ''}${a.requisitante ? `<br><span class="apoio" style="font-size:12px">gestor ${esc(a.requisitante)}</span>` : ''}</td><td class="num">${prazoChamadoTi(a) ? dm(prazoChamadoTi(a)) + (chamadoTiAtrasado(a) ? ' ' + chip('c-cr', 'cr', 'atrasado') : '') : '—'}</td><td>${situacao(a)}</td><td>${a.chamado_ti ? `<b>${esc(a.chamado_ti)}</b>` : souTi() ? `<span class="linha-flex" style="gap:4px"><input class="sel-doc" id="tiCh-${a.id}" placeholder="nº do chamado" style="width:120px" aria-label="Número do chamado"><button class="bt peq" data-act="tiChamado" data-id="${a.id}">Salvar</button></span>` : '—'}</td></tr>`;
  return `<div class="cab"><div class="t"><span class="olho">TI</span><h1>Equipamentos e acessos</h1><p class="apoio">Quem vai entrar e precisa de equipamento ou acesso, desde o cadastro pelo R&S (previsão) até a matrícula (pedido). O prazo é o do local: ${esc(String((E.locais[0] || {}).prazo_ti_horas || 48))} horas úteis antes do primeiro dia, 72 nos CDs. Registre aqui o número do chamado; a Admissão vê na ficha. Cancelamentos aparecem embaixo para desligar acessos e recolher equipamento.</p></div></div>
   <div class="grade g4"><div class="caixa kpi"><b class="num">${l.length}</b><span>pessoas com equipamento</span></div><div class="caixa kpi"><b class="num">${em7}</b><span>entram em 7 dias</span></div><div class="caixa kpi"><b class="num">${semChamado}</b><span>com matrícula e sem chamado</span></div><div class="caixa kpi"><b class="num">${atras}</b><span>fora do prazo</span></div></div>
   <div class="caixa pilha"><h2>Próximas entradas</h2>${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Integração</th><th>Equipamentos e acessos</th><th>Prazo do chamado</th><th>Situação</th><th>Chamado</th></tr></thead><tbody>${l.map(linha).join('')}</tbody></table></div>` : '<div class="vazio">Ninguém com equipamento previsto.</div>'}</div>
   <div class="caixa pilha"><h2>Admissões canceladas (últimos 30 dias)</h2><p class="apoio" style="font-size:12.5px">Cancelar o chamado, desligar os acessos e recolher o equipamento reservado.</p>${canc.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Era para</th><th>Equipamentos</th><th>Chamado</th><th>Cancelada em</th></tr></thead><tbody>${canc.map((a) => `<tr><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.cargo)} · ${esc(a.local)}</span></td><td class="num">${dmd(a.data_inicio)}</td><td style="font-size:12.5px">${esc((a.ti_itens || []).join(', '))}</td><td>${esc(a.chamado_ti || '—')}</td><td class="num">${quando(a.declinou_em)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nenhum cancelamento recente.</div>'}</div>`;
};
VIEWS.entregas = () => {
  const l = E.entregas;
  return `<div class="cab"><div class="t"><span class="olho">Especialista</span><h1>Minhas entregas</h1><p class="apoio">Só aparece quem está com a documentação completa e precisa de algo da sua área. Marque quando estiver pronto: a Admissão vê sem precisar de e-mail de volta.</p></div></div>
  ${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Área</th><th>Integração</th><th>Detalhe</th><th>O que preparar</th><th>Status</th></tr></thead><tbody>${l.map((x) => `<tr><td><b>${esc(x.nome)}</b><br><span class="apoio" style="font-size:12px">${esc(x.cargo)} · ${esc(x.local)}</span></td><td>${esc(x.area_nome)}</td><td class="num">${dmd(x.data_inicio)} ${esc(x.hora_integ || '')}<br><span class="apoio" style="font-size:12px">${esc(x.integ || '')}</span></td><td style="font-size:12.5px">${esc([x.matricula && 'Matrícula ' + x.matricula, x.camisa && 'Camisa ' + x.camisa, x.fretado && 'Fretado ' + x.fretado, (x.ti_itens || []).length && 'TI: ' + x.ti_itens.join(', '), x.usa_frota && 'Usa frota', x.pcd && 'PCD: sim', x.requisitante && 'Gestor ' + x.requisitante].filter(Boolean).join(' · '))}</td><td style="font-size:12.5px">${esc(x.leva)}</td><td>${x.pronto_em ? chip('c-ok', 'ok', 'Pronto ' + dm(new Date(x.pronto_em))) : `<button class="bt peq" data-act="entregar" data-id="${x.id}" data-area="${esc(x.area_id)}">Marcar pronto</button>`}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nada para preparar agora.</div>'}`;
};

// --------------------------------------------------------------- turma
// Lista da turma no layout U1/U2 (16 colunas), xlsx/csv e envio aos destinatários da unidade.
const COLS_LISTA = [['local', 'LOCAL'], ['data_admissao', 'DATA DE ADMISSÃO'], ['horario_agendado', 'HORÁRIO AGENDADO'], ['analista_responsavel', 'ANALISTA RESPONSÁVEL'], ['requisitante', 'REQUISITANTE'], ['cargo', 'CARGO'], ['matricula', 'MATRICULA'], ['nome', 'NOME'], ['telefone', 'TELEFONE'], ['pcd', 'PCD'], ['tipo', 'TIPO'], ['chamado', 'CHAMADO'], ['dias_integracao', 'DIAS INTEGRAÇÃO'], ['integracao_nr12', 'INTEGRAÇÃO NR12'], ['cc_codigo', 'Código Centro de Custo'], ['cc_nome', 'Nome Centro de Custo']];
const datasTurma = () => [...new Set([...turmas(), ...E.turmasCal.filter((t) => t.ativo).map((t) => t.data)])].sort();
function matrizLista(l) {
  const fmtData = (d) => (d ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : '');
  return [COLS_LISTA.map((c) => c[1]), ...l.map((x) => COLS_LISTA.map(([k]) => (k === 'data_admissao' ? fmtData(x[k]) : x[k] == null ? '' : x[k])))];
}
function nomeLista(data, locais) {
  const rot = locais.map((n) => localDe(n).rotulo_lista || n.toUpperCase()).join(' e ');
  return nomeArquivoSeguro('INTEGRACAO ' + rot + ' ' + data.split('-').reverse().join('-'));
}
function planilhaLista(l) {
  if (!window.XLSX) throw new Error('O gerador de Excel não carregou. Verifique a internet.');
  const ws = XLSX.utils.aoa_to_sheet(matrizLista(l)); ws['!cols'] = COLS_LISTA.map(([k]) => ({ wch: k === 'nome' ? 34 : k === 'cargo' || k === 'cc_nome' ? 26 : 16 }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Admitidos');
  return wb;
}
VIEWS.turma = () => {
  const ds = datasTurma(); if (!E.turmaLista || !ds.includes(E.turmaLista)) E.turmaLista = ds.find((d) => d >= iso(new Date())) || ds[ds.length - 1] || null;
  const doDia = vivas().filter((a) => a.data_inicio === E.turmaLista);
  const locaisDia = [...new Set(doDia.map((a) => a.local))].sort();
  if (!E.locaisLista || !E.locaisLista.length || E.locaisLista.some((x) => !locaisDia.includes(x))) E.locaisLista = locaisDia.slice();
  const cal = E.turmasCal.filter((t) => t.ativo && t.data === E.turmaLista);
  const lista = doDia.filter((a) => E.locaisLista.includes(a.local) && (!E.soMatriculados || a.matricula)).sort((x, y) => (x.local + x.cargo + x.nome).localeCompare(y.local + y.cargo + y.nome))
    .map((a) => ({ id: a.id, local: a.local, data_admissao: a.data_inicio, horario_agendado: a.hora_integ, analista_responsavel: a.analista_rs, requisitante: a.requisitante, cargo: a.cargo, matricula: a.matricula, nome: a.nome, telefone: a.telefone, pcd: a.pcd === 'Não' ? 'NÃO' : 'SIM', tipo: (a.ti_itens || []).join(', ') || null, chamado: a.chamado_ti, dias_integracao: a.dias_integ, integracao_nr12: a.nr12 ? 'SIM' : 'NÃO', cc_codigo: a.cc_codigo, cc_nome: a.cc_nome, requisitante_email: a.requisitante_email }));
  E.listaAtual = lista;
  const semMat = doDia.filter((a) => E.locaisLista.includes(a.local) && !a.matricula).length;
  const destinos = (() => { const ls = E.locaisLista.map((n) => localDe(n)); const s = new Set(); E.dest.filter((d) => d.ativo && (d.unidade === 'todas' || E.locaisLista.includes(d.unidade) || ls.some((l) => l.tipo === d.unidade || l.grupo_turma === d.unidade))).forEach((d) => (d.emails || []).forEach((e) => s.add(e))); return [...s]; })();
  const reqs = [...new Set(lista.map((x) => x.requisitante_email).filter(Boolean))];
  return `<div class="cab"><div class="t"><span class="olho">Integração</span><h1>Turma e lista de admitidos</h1><p class="apoio">A lista que vai na sexta às 12h para as áreas, no layout U1/U2 (16 colunas), montada do que está na plataforma. Baixe em Excel ou CSV, ou envie direto para os destinatários cadastrados por unidade, com os requisitantes em cópia. O convite ao candidato sai sozinho quando a matrícula é registrada${ligado('convite_ativo') ? '' : ' (desligado no parâmetro convite_ativo)'}.</p></div></div>
  <div class="filtros"><label class="olho" for="tlData">Turma</label><select id="tlData">${ds.map((d) => `<option value="${d}" ${d === E.turmaLista ? 'selected' : ''}>${dmd(d)}${E.turmasCal.some((t) => t.data === d) ? '' : ' · fora do cronograma'}</option>`).join('')}</select>
    <span class="olho">Unidades</span><div class="marcas">${locaisDia.map((n) => `<label><input type="checkbox" data-loc="${esc(n)}" ${E.locaisLista.includes(n) ? 'checked' : ''}> ${esc(n)}</label>`).join('') || '<span class="apoio">ninguém nesta data</span>'}</div>
    <label class="marcas"><label><input type="checkbox" id="tlSo" ${E.soMatriculados ? 'checked' : ''}> Só matriculados</label></label></div>
  ${cal.length ? cal.map((t) => `<div class="faixa"><p><b>${esc(t.nome || GRUPO_TURMA[t.grupo])}</b> · inclusão até ${dmh(t.corte_inclusao)}${t.corte_remuneracao ? ' · Remuneração até ' + dmh(t.corte_remuneracao) : ''} · documentos, ASO e LG até ${dmh(t.corte_documentos)}${t.formato ? ' · ' + (t.formato === 'online' ? 'on-line' : 'presencial') : ''}${t.link_ou_endereco ? ' · ' + esc(t.link_ou_endereco) : ''}${t.contato_chegada ? ' · procurar ' + esc(t.contato_chegada) : ''}</p></div>`).join('') : (E.turmaLista ? '<div class="aviso-faixa">Esta data não está no cronograma de turmas. Cadastre em "Cronograma de turmas", abaixo, para o corte e o convite usarem os horários certos.</div>' : '')}
  <div class="caixa pilha"><div class="linha-flex"><h2 style="flex:1">${plural(lista.length, 'pessoa', 'pessoas')} na lista${semMat && E.soMatriculados ? ` <span class="apoio" style="font-size:13px">· ${plural(semMat, 'ainda sem matrícula', 'ainda sem matrícula')}</span>` : ''}</h2>
    <button class="bt" data-act="listaXlsx" ${lista.length ? '' : 'disabled'}>Baixar .xlsx</button><button class="bt" data-act="listaCsv" ${lista.length ? '' : 'disabled'}>Baixar .csv</button>
    ${tem('admissao') ? `<button class="bt pri" data-act="listaEnviar" ${lista.length && E.soMatriculados ? '' : 'disabled'} title="${E.soMatriculados ? '' : 'Marque só matriculados para enviar'}">Enviar por e-mail</button>` : ''}</div>
    ${tem('admissao') ? `<p class="apoio" style="font-size:12.5px">Vai para: ${destinos.length ? esc(destinos.join(', ')) : '<b>ninguém cadastrado para estas unidades</b>'}${reqs.length ? ' · em cópia os requisitantes: ' + esc(reqs.join(', ')) : ''}. Cadastre e ajuste em "Especialistas e avisos".</p>` : ''}
    ${lista.length ? `<div class="tab-wrap"><table class="t-lista"><thead><tr>${COLS_LISTA.map((c) => `<th>${esc(c[1])}</th>`).join('')}</tr></thead><tbody>${lista.map((x) => `<tr class="clic" data-act="ficha" data-id="${x.id}">${COLS_LISTA.map(([k]) => `<td>${esc(k === 'data_admissao' ? dm(x[k]) : x[k] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Ninguém para listar com estes filtros.</div>'}</div>
  ${tem('admissao') ? cronogramaTurmas() : ''}`;
};
POS.turma = () => {
  $('#tlData')?.addEventListener('change', (e) => { E.turmaLista = e.target.value; E.locaisLista = null; render(); });
  $('#tlSo')?.addEventListener('change', (e) => { E.soMatriculados = e.target.checked; render(); });
  document.querySelectorAll('[data-loc]').forEach((c) => c.addEventListener('change', () => { E.locaisLista = [...document.querySelectorAll('[data-loc]:checked')].map((x) => x.dataset.loc); render(); }));
  $('#fTurmaNova')?.addEventListener('submit', salvarTurmaForm);
  $('#tnData')?.addEventListener('change', sugerirCortes); $('#tnGrupo')?.addEventListener('change', sugerirCortes);
};
// cronograma de turmas (DP): data, grupo e os três cortes; a regra sugere os horários
function cronogramaTurmas() {
  const hoje = iso(new Date());
  const ts = E.turmasCal.filter((t) => t.data >= iso(new Date(Date.now() - 30 * DIA))).sort((a, b) => a.data.localeCompare(b.data) || a.grupo.localeCompare(b.grupo));
  const ed = E.turmaEdit || {};
  return `<div class="caixa pilha" id="cronograma"><h2>Cronograma de turmas</h2><p class="apoio" style="font-size:13px">O cronograma mensal do DP, dentro da plataforma: o cadastro escolhe a turma, o contador de corte e as cobranças usam estes horários, e o convite usa o formato, o link ou endereço e quem procurar. Regra sugerida: inclusão até 7 dias antes ao meio-dia, Remuneração no mesmo dia até as 18h, documentos, ASO e LG até 4 dias antes ao meio-dia.</p>
    <form id="fTurmaNova" class="form" style="grid-template-columns:repeat(6,minmax(0,1fr))">
      <input type="hidden" id="tnId" value="${esc(ed.id || '')}">
      <div class="campo"><label for="tnData">Data de admissão</label><input id="tnData" type="date" value="${esc(ed.data || '')}" required></div>
      <div class="campo"><label for="tnGrupo">Unidades</label><select id="tnGrupo">${Object.entries(GRUPO_TURMA).map(([k, v]) => `<option value="${k}" ${k === (ed.grupo || 'pa_sp_comercial') ? 'selected' : ''}>${v}</option>`).join('')}</select></div>
      <div class="campo"><label for="tnInc">Inclusão até</label><input id="tnInc" type="datetime-local" value="${paraLocal(ed.corte_inclusao)}" required></div>
      <div class="campo"><label for="tnRem">Remuneração até</label><input id="tnRem" type="datetime-local" value="${paraLocal(ed.corte_remuneracao)}"></div>
      <div class="campo"><label for="tnDoc">Documentos até</label><input id="tnDoc" type="datetime-local" value="${paraLocal(ed.corte_documentos)}" required></div>
      <div class="campo"><label for="tnFormato">Formato</label><select id="tnFormato"><option value="">Padrão do local</option><option value="presencial" ${ed.formato === 'presencial' ? 'selected' : ''}>Presencial</option><option value="online" ${ed.formato === 'online' ? 'selected' : ''}>On-line</option></select></div>
      <div class="campo l2"><label for="tnLink">Link do Teams ou endereço</label><input id="tnLink" value="${esc(ed.link_ou_endereco || '')}"></div>
      <div class="campo l2"><label for="tnContato">Quem procurar ao chegar</label><input id="tnContato" value="${esc(ed.contato_chegada || '')}"></div>
      <div class="campo"><label for="tnObs">Observação</label><input id="tnObs" value="${esc(ed.observacao || '')}" placeholder="feriado, mutirão…"></div>
      <div class="campo" style="justify-content:flex-end"><div class="linha-flex"><button class="bt pri" type="submit">${ed.id ? 'Salvar turma' : 'Incluir turma'}</button>${ed.id ? '<button class="bt" type="button" data-act="turmaCancelar">Cancelar</button>' : ''}</div></div>
    </form>
    ${ts.length ? `<div class="tab-wrap"><table><thead><tr><th>Admissão</th><th>Unidades</th><th>Inclusão até</th><th>Remuneração até</th><th>Documentos até</th><th>Formato · onde</th><th>Pessoas</th><th></th></tr></thead><tbody>${ts.map((t) => `<tr ${t.data < hoje ? 'style="opacity:.6"' : ''}><td class="num"><b>${dmd(t.data)}</b>${t.ativo ? '' : ' ' + chip('c-ag', '', 'inativa')}</td><td>${esc(t.nome || GRUPO_TURMA[t.grupo])}</td><td class="num">${dmh(t.corte_inclusao)}</td><td class="num">${t.corte_remuneracao ? dmh(t.corte_remuneracao) : '—'}</td><td class="num">${dmh(t.corte_documentos)}</td><td style="font-size:12.5px">${t.formato ? (t.formato === 'online' ? 'on-line' : 'presencial') : 'padrão'}${t.link_ou_endereco ? ' · ' + esc(t.link_ou_endereco) : ''}${t.contato_chegada ? ' · ' + esc(t.contato_chegada) : ''}</td><td class="num">${vivas().filter((a) => a.turma_id === t.id).length}</td><td><div class="linha-flex"><button class="bt peq" data-act="turmaEditar" data-id="${t.id}">Editar</button>${vivas().some((a) => a.turma_id === t.id) ? '' : `<button class="bt peq fan" data-act="turmaApagar" data-id="${t.id}">Apagar</button>`}</div></td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nenhuma turma cadastrada. Inclua as datas do cronograma do mês.</div>'}</div>`;
}
function sugerirCortes() {
  const d = $('#tnData')?.value; if (!d) return;
  const base = deIso(d); const ajusta = (dias, h) => { const x = new Date(base.getTime() - dias * DIA); x.setHours(h, 0, 0, 0); return paraLocal(x.toISOString()); };
  if (!$('#tnInc').value) $('#tnInc').value = ajusta(7, 12);
  if (!$('#tnRem').value) $('#tnRem').value = ajusta(7, 18);
  if (!$('#tnDoc').value) $('#tnDoc').value = ajusta(4, 12);
}
async function salvarTurmaForm(e) {
  e.preventDefault();
  const v = (id) => $('#' + id).value.trim();
  if (!v('tnData') || !v('tnInc') || !v('tnDoc')) { toast('Data, inclusão e documentos são obrigatórios.', true); return; }
  const o = { data: v('tnData'), grupo: v('tnGrupo'), nome: GRUPO_TURMA[v('tnGrupo')], corte_inclusao: deLocal(v('tnInc')), corte_remuneracao: deLocal(v('tnRem')), corte_documentos: deLocal(v('tnDoc')), formato: v('tnFormato') || null, link_ou_endereco: v('tnLink') || null, contato_chegada: v('tnContato') || null, observacao: v('tnObs') || null, ativo: true };
  if (v('tnId')) o.id = v('tnId');
  if (await tentar(() => api.salvarTurma(o), o.id ? 'Turma salva' : 'Turma incluída')) { E.turmaEdit = null; await recarregar(); }
}

// --------------------------------------------------------- especialistas
VIEWS.especialistas = () => {
  const g = Object.keys(GATILHO);
  const unidades = ['todas', ...E.locais.map((l) => l.nome), 'fabrica', 'cd', 'sp', 'comercial', 'pa_sp_comercial', 'ssbv_cds'];
  const optU = (sel) => unidades.map((u) => `<option value="${esc(u)}" ${u === sel ? 'selected' : ''}>${esc(GRUPO_TURMA[u] || ({ todas: 'Todas as unidades', fabrica: 'Fábricas', cd: 'CDs', sp: 'Escritórios SP', comercial: 'Comercial' })[u] || u)}</option>`).join('');
  const linhaBp = (b, i) => `<div class="area" data-bp="${i}"><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(b.unidade)}</select></div><div class="ac"><label class="olho">Setor (opcional)</label><input class="sel-doc" data-f="area" value="${esc(b.area || '')}" placeholder="vazio = unidade inteira"></div><div class="ac"><label class="olho">BP</label><input class="sel-doc" data-f="nome" value="${esc(b.nome || '')}" placeholder="Nome"><input class="sel-doc" data-f="email" value="${esc(b.email || '')}" placeholder="nome@grupocimed.com.br" style="margin-top:4px"></div><div class="linha-flex"><button class="bt peq" data-act="bpSalvar" data-i="${i}">Salvar</button>${b.id ? `<button class="bt peq fan" data-act="bpApagar" data-id="${b.id}">Apagar</button>` : ''}</div></div>`;
  const linhaDest = (d, i) => `<div class="area" data-dest="${i}"><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(d.unidade)}</select><input class="sel-doc" data-f="descricao" value="${esc(d.descricao || '')}" placeholder="Descrição" style="margin-top:4px"></div><div class="ac" style="grid-column:span 2"><label class="olho">E-mails e listas</label><input class="sel-doc" style="width:100%;min-height:40px" data-f="emails" value="${esc((d.emails || []).join(', '))}" placeholder="lista@grupocimed.com.br, portaria@grupocimed.com.br"><label class="marcas" style="margin-top:4px"><label><input type="checkbox" data-f="copiar_requisitantes" ${d.copiar_requisitantes !== false ? 'checked' : ''}> Requisitantes da turma em cópia</label></label></div><div class="linha-flex"><button class="bt peq" data-act="destSalvar" data-i="${i}">Salvar</button>${d.id ? `<button class="bt peq fan" data-act="destApagar" data-id="${d.id}">Apagar</button>` : ''}</div></div>`;
  return `<div class="cab"><div class="t"><span class="olho">Comunicação</span><h1>Especialistas e avisos</h1><p class="apoio">Quem recebe e-mail, em que momento e com quais dados. Área sem e-mail cadastrado não recebe nada. As marcadas como sugestão vieram do mapeamento e precisam da sua validação antes de ligar. Separe vários e-mails por vírgula. Gestor, analista de R&S e BP têm destino por candidato: o e-mail do requisitante na ficha, o e-mail da analista no cadastro de pessoas e a tabela de BPs abaixo; a lista da área é só reserva.</p></div></div>
  ${g.map((k) => `<div class="caixa" style="padding:0"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>${esc(GATILHO[k])}</h2></div>
   ${E.areas.filter((a) => a.gatilho === k).map((a) => `<div class="area"><div><b>${esc(a.nome)}</b><div class="linha-flex" style="gap:6px;margin-top:3px">${a.origem === 'sugestao' ? chip('c-ama', '', 'Sugestão') : chip('c-ag', 'ok', 'Você pediu')}<span class="apoio" style="font-size:12px">${esc(CONDICAO[a.condicao] || a.condicao)}</span></div><div class="apoio" style="font-size:12.5px;margin-top:4px">${esc(a.leva)}</div></div>
     <div class="ac" style="grid-column:span 2"><label class="olho" for="em_${esc(a.id)}">E-mails${['gestor', 'rs', 'bp'].includes(a.id) ? ' (reserva)' : ''}</label><input id="em_${esc(a.id)}" class="sel-doc" style="width:100%;min-height:40px" value="${esc((a.emails || []).join(', '))}" placeholder="nome@grupocimed.com.br"></div>
     <div class="linha-flex"><button class="troca" data-act="togArea" data-id="${esc(a.id)}" aria-pressed="${a.ativo}" aria-label="Ligar ${esc(a.nome)}"></button><button class="bt peq" data-act="salvarArea" data-id="${esc(a.id)}">Salvar</button></div></div>`).join('')}</div>`).join('')}
  <div class="caixa" style="padding:0" id="bpsBox"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>BPs por unidade</h2><p class="apoio" style="font-size:12.5px">Quem recebe o aviso "BP da unidade" quando a documentação fica completa. Uma linha por unidade; com setor preenchido, vale só para aquele setor.</p></div>
   ${E.bps.map(linhaBp).join('')}${linhaBp({ unidade: 'todas' }, E.bps.length)}</div>
  <div class="caixa" style="padding:0" id="destBox"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>Lista da turma: destinatários por unidade</h2><p class="apoio" style="font-size:12.5px">Quem recebe o e-mail "INTEGRAÇÃO SITE I dd/mm/aaaa" com a planilha, enviado pela tela Turma e lista. Listas de distribuição valem.</p></div>
   ${E.dest.map(linhaDest).join('')}${linhaDest({ unidade: 'fabrica' }, E.dest.length)}</div>
  <div class="caixa pilha"><p class="apoio">Tipo de deficiência e CPF só vão para Saúde, Admissão e Remuneração. As demais áreas recebem "PCD: sim", por ser dado sensível pela LGPD. As respostas do Termo de Transparência só o papel Compliance abre.</p></div>`;
};

// ------------------------------------------------------------ compliance
VIEWS.compliance = () => {
  const l = E.termos; const t = l.find((x) => x.id === E.selTermo);
  const linha = (x) => `<button class="item-esq" data-act="selTermo" data-id="${x.id}" aria-current="${x.id === E.selTermo}"><span class="linha-flex" style="gap:6px">${x.algum_sim ? chip('c-at', 'at', 'Há "Sim"') : chip('c-ok', 'ok', 'Tudo "Não"')}${x.tratado_em ? chip('c-ag', 'ok', 'Tratado') : ''}<small class="num">${quando(x.assinado_em || x.criado_em)}</small></span><b>${esc(titulo(x.nome))}</b><small>${esc(x.cargo || '')}${x.local ? ' · ' + esc(x.local) : ''}${x.matricula ? ' · ' + esc(x.matricula) : ''}</small></button>`;
  const resp = t ? t.respostas || {} : {};
  return `<div class="cab"><div class="t"><span class="olho">Compliance</span><h1>Termo de Transparência</h1><p class="apoio">Os termos assinados no kit admissional, com as respostas às 11 perguntas. Só quem tem o papel Compliance vê esta tela e estas respostas: nem o R&S, nem a Admissão, nem o gestor. Qualquer "Sim" gera um e-mail para a área Compliance; registre aqui o parecer quando a Comissão de Ética e Conduta concluir.</p></div></div>
  ${l.length ? `<div class="grade" style="grid-template-columns:minmax(0,360px) minmax(0,1fr)"><div class="lista-esq">${l.map(linha).join('')}</div>
   ${t ? `<article class="caixa pilha"><div><span class="olho">Termo assinado em ${quando(t.assinado_em || t.criado_em)}${t.codigo ? ' · código ' + esc(t.codigo) : ''} · ${t.origem === 'kit_online' ? 'celular do colaborador' : 'computador de apoio da integração'}</span><h2>${esc(titulo(t.nome))}</h2><span class="apoio">${esc(t.cargo || '—')}${t.setor ? ' · ' + esc(t.setor) : ''} · ${esc(t.local || '—')}${t.matricula ? ' · matrícula ' + esc(t.matricula) : ''}${t.cpf ? ' · CPF ' + fmtCpf(t.cpf) : ''}</span></div>
     <div class="docs">${TERMO_PERGUNTAS.map(([k, p, d]) => { const v = resp[k]; const sim = String(v || '').toLowerCase() === 'sim'; return `<div class="doc" ${sim ? 'style="border-color:var(--at)"' : ''}><div><div class="dn">${k.slice(1)}. ${esc(p)}</div>${sim && resp[k + '_det'] ? `<div class="dm"><b>${esc(resp[k + '_det'])}</b></div>` : sim ? `<div class="motivo">Sem detalhe informado (${esc(d)})</div>` : ''}</div><div class="acoes">${v == null || v === '' ? chip('c-ag', 'ag', 'Sem resposta') : sim ? chip('c-at', 'at', 'Sim') : chip('c-ok', 'ok', 'Não')}</div></div>`; }).join('')}</div>
     <div class="apoio" style="font-size:12.5px">Código de Conduta: ${resp.codigo_conduta_aberto_em ? 'link aberto em ' + quando(resp.codigo_conduta_aberto_em) : 'link não foi aberto na tela'}${resp.codigo_conduta_ciente ? ' · declarou ciência' : ''}.</div>
     ${t.tratado_em ? `<div class="faixa"><p><b>Tratado</b> em ${quando(t.tratado_em)} por ${esc(t.tratado_por || '')}.${t.parecer ? ' Parecer: ' + esc(t.parecer) : ''}</p></div>` : `<form id="fParecer" class="linha-flex"><input id="parecerTx" maxlength="2000" placeholder="Parecer da Comissão de Ética (fica com seu nome e a data)" style="flex:1;min-width:200px;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup)"><button class="bt pri" type="submit">Marcar como tratado</button></form>`}
   </article>` : ''}</div>` : '<div class="vazio">Nenhum termo assinado ainda. Eles chegam quando o colaborador assina o kit na integração.</div>'}`;
};
POS.compliance = () => {
  $('#fParecer')?.addEventListener('submit', async (e) => { e.preventDefault(); const tx = $('#parecerTx').value.trim(); if (await tentar(() => api.termoTratar(E.selTermo, tx), 'Termo marcado como tratado')) await recarregar(); });
};

// ---------------------------------------------------------------- e-mails
VIEWS.emails = () => {
  const l = E.emails; if (!E.selMail || !l.some((m) => m.id === E.selMail)) E.selMail = l[0]?.id;
  const m = l.find((x) => x.id === E.selMail);
  const stc = { pendente: chip('c-ag', 'ag', 'Na fila'), enviando: chip('c-info', 'ag', 'Enviando'), enviado: chip('c-ok', 'ok', 'Enviado'), erro: chip('c-cr', 'cr', 'Erro'), descartado: chip('c-ag', 'cr', 'Descartado') };
  const tipo = TIPO_EMAIL;
  return `<div class="cab"><div class="t"><span class="olho">Comunicação</span><h1>Fila de e-mails</h1><p class="apoio">Tudo o que a plataforma manda sozinha, com o texto exato. Os e-mails saem da caixa ${esc(E.params.remetente || '')}.</p></div></div>
  ${E.params.envio_ativo !== 'sim' ? `<div class="aviso-faixa">${ICO.at.replace('<svg', '<svg width="16" height="16"')}Modo homologação: os e-mails ficam na fila e nada é enviado. Para ligar, mude o parâmetro envio_ativo para "sim" no Supabase.</div>` : ''}
  ${E.params.redirecionar_para ? `<div class="aviso-faixa info">${ICO.info.replace('<svg', '<svg width="16" height="16"')}Modo teste: todos os e-mails vão só para ${esc(E.params.redirecionar_para)}.</div>` : ''}
  <details class="caixa" style="margin-bottom:12px"><summary style="cursor:pointer;font-weight:700">Jornada de e-mails do colaborador (da aprovação em diante)</summary>
   <div class="docs" style="margin-top:10px">${JORNADA_EMAILS.map(([t, q, d], i) => `<div class="doc"><div><div class="dn">${i + 1}. ${esc(TIPO_EMAIL[t] || t)} <span class="apoio" style="font-weight:400">· quando: ${esc(q)}</span></div><div class="dm">${esc(d)}</div></div><div class="acoes">${chip(E.emails.some((m) => m.tipo === t) ? 'c-ok' : 'c-ag', E.emails.some((m) => m.tipo === t) ? 'ok' : '', plural(E.emails.filter((m) => m.tipo === t).length, 'na fila ou enviado', 'na fila ou enviados'))}</div></div>`).join('')}</div>
   <p class="apoio" style="font-size:12.5px;margin:10px 0 4px">Quando acontece algo:</p>
   <div class="docs">${EMAILS_EVENTO.map(([t, q, d]) => `<div class="doc"><div><div class="dn">${esc(TIPO_EMAIL[t] || t)} <span class="apoio" style="font-weight:400">· ${esc(q)}</span></div><div class="dm">${esc(d)}</div></div></div>`).join('')}</div>
   <p class="apoio" style="font-size:12.5px;margin:10px 0 4px">Para a equipe e as áreas (não vão ao colaborador):</p>
   <div class="docs">${EMAILS_EQUIPE.map(([t, q, d]) => `<div class="doc"><div><div class="dn">${esc(TIPO_EMAIL[t] || t)} <span class="apoio" style="font-weight:400">· ${esc(q)}</span></div><div class="dm">${esc(d)}</div></div></div>`).join('')}</div>
   ${E.params.whatsapp_ativo === 'sim' ? '' : `<p class="apoio" style="font-size:12.5px;margin-top:8px">WhatsApp: o envio automático está desligado (parâmetro whatsapp_ativo). O número do time (${esc(E.params.whatsapp_admissao || '')}) aparece nos e-mails, e a Esteira abre a conversa com o texto pronto (link do kit ou PDF). Para ligar é preciso a API oficial do WhatsApp Business: README, seção 15.</p>`}
  </details>
  ${E.params.whatsapp_ativo === 'sim' ? secaoWhatsapp() : ''}
  <div class="grade" style="grid-template-columns:minmax(0,360px) minmax(0,1fr)">
   <div class="mail-lista">${l.map((x) => `<button class="mail" data-act="selMail" data-id="${x.id}" aria-current="${x.id === E.selMail}"><span class="linha-flex" style="gap:6px">${stc[x.status] || ''}${chip('c-info', '', tipo[x.tipo] || x.tipo)}<small class="num">${quando(x.criado_em)}</small></span><span class="as">${esc(x.assunto)}</span><small>Para: ${esc(x.para_nome || (x.para || []).join(', '))}</small></button>`).join('') || '<div class="vazio">Nenhum e-mail ainda.</div>'}</div>
   ${m ? `<article class="carta"><div class="faixa-mail"></div><div class="hdr"><b>Para</b><span>${esc((m.para || []).join(', '))}</span><b>Assunto</b><span><b style="color:var(--tx)">${esc(m.assunto)}</b></span><b>Situação</b><span>${stc[m.status] || ''} ${m.enviado_em ? 'em ' + quando(m.enviado_em) : ''}${m.erro ? '<br><span style="color:var(--cr)">' + esc(m.erro) + '</span>' : ''}</span>${Array.isArray(m.anexos) && m.anexos.length ? `<b>Anexo</b><span>${esc(m.anexos.map((x) => x.nome).join(', '))}</span>` : ''}</div><div class="corpo">${esc(m.corpo)}</div>${['erro', 'descartado'].includes(m.status) ? `<div style="padding:0 18px 18px"><button class="bt" data-act="reenviar" data-id="${m.id}">Tentar de novo</button></div>` : ''}</article>` : ''}
  </div>`;
};

// WhatsApp (SQL 22): o link do kit on-line vai pela API oficial da Meta. A seção só aparece
// com whatsapp_ativo = sim; cada linha mostra a situação do envio, como a fila de e-mails.
function secaoWhatsapp() {
  const l = E.whatsapp || [];
  const stc = { pendente: chip('c-ag', 'ag', 'Na fila'), enviando: chip('c-info', 'ag', 'Enviando'), enviado: chip('c-ok', 'ok', 'Enviado'), erro: chip('c-cr', 'cr', 'Erro'), descartado: chip('c-ag', 'cr', 'Descartado') };
  const nomeDe = (w) => { const a = E.admissoes.find((x) => x.id === w.admissao_id); return a ? a.nome : (w.parametros && w.parametros.nome) || ''; };
  const fone = (n) => { const d = String(n || '').replace(/\D/g, '').replace(/^55/, ''); return d.length >= 10 ? `(${d.slice(0, 2)}) ${d.slice(2, -4)}-${d.slice(-4)}` : n || ''; };
  const desvio = String(E.params.whatsapp_redirecionar_para || '').trim();
  return `<div class="caixa" style="margin-bottom:12px"><div class="linha-flex" style="justify-content:space-between;align-items:baseline"><b>WhatsApp · link do kit on-line</b><span class="apoio">${plural(l.filter((w) => w.status === 'pendente').length, 'na fila', 'na fila')} · ${plural(l.filter((w) => w.status === 'erro').length, 'com erro', 'com erro')}</span></div>
   <p class="apoio" style="font-size:12.5px;margin:4px 0 8px">Quando a Esteira publica o kit on-line, o link vai para o WhatsApp do colaborador pelo modelo "${esc(E.params.whatsapp_modelo_kit || 'kit_admissao_link')}" da API oficial. A mensagem sai em até 5 minutos.${desvio ? ` <b>Modo teste:</b> todas as mensagens vão só para ${esc(fone(desvio))}.` : ''}</p>
   ${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Quando</th><th>Pessoa</th><th>Número</th><th>Situação</th><th></th></tr></thead><tbody>${l.map((w) => `<tr><td class="num">${quando(w.criado_em)}</td><td>${esc(nomeDe(w))}</td><td class="num">${esc(fone(w.numero))}</td><td>${stc[w.status] || esc(w.status)}${w.enviado_em ? ` <small class="apoio">em ${quando(w.enviado_em)}</small>` : ''}${w.erro ? `<br><small style="color:var(--cr)">${esc(w.erro)}</small>` : ''}${w.tentativas > 1 ? ` <small class="apoio">(${w.tentativas} tentativas)</small>` : ''}</td><td>${['erro', 'descartado'].includes(w.status) ? `<button class="bt peq" data-act="reenviarZap" data-id="${w.id}">Tentar de novo</button>` : ''}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nenhuma mensagem de WhatsApp ainda.</div>'}
  </div>`;
}

// ============================================== IMPORTAR A PLANILHA
async function lerPlanilha(file) {
  if (!window.XLSX) { toast('O leitor de Excel não carregou. Verifique a internet.', true); return; }
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const aba = wb.SheetNames.find((n) => /RELA..O ADMITIDOS/i.test(n)) || wb.SheetNames[0];
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[aba], { header: 1, raw: true, defval: null });
  const hi = linhas.findIndex((r) => r && r.some((v) => /NOME CANDIDATO/i.test(String(v || ''))));
  if (hi < 0) { toast('Não achei a coluna "NOME CANDIDATO" na planilha.', true); return; }
  const H = linhas[hi].map((v) => String(v || '').trim().toUpperCase());
  const col = (re) => H.findIndex((h) => re.test(h));
  const C = { ini: col(/DATA INICIO/), nome: col(/NOME CANDIDATO/), cpf: col(/^CPF/), integ: col(/^INTEGRA/), hora: col(/HOR.RIO INTEGRA/), dias: col(/DIAS INTEGRA/), cargo: col(/^CARGO/), setor: col(/^SETOR/), local: col(/^LOCAL/), escala: col(/^ESCALA/), pos: col(/^POSI/), proc: col(/PROCESSO/), req: col(/REQUISITANTE/), an: col(/ANALISTA/), tel: col(/TELEFONE/), email: col(/E-MAIL/), camisa: col(/CAMISA/), pcd: col(/PCD/), fret: col(/FRETADO/), mat: col(/MATR/), doc: col(/DOCUMENTA/), dp: col(/ADM PESSOAL/), cham: col(/CHAMADO/), rem: col(/REMUNERA/), exame: col(/EXAME M/) };
  const g = (r, k) => (C[k] >= 0 ? r[C[k]] : null); const tx = (v) => (v == null ? '' : String(v).trim());
  const data = (v) => (v instanceof Date ? iso(v) : typeof v === 'number' ? iso(new Date(Math.round((v - 25569) * DIA) + new Date().getTimezoneOffset() * 6e4)) : '');
  const hora = (v) => (v instanceof Date ? String(v.getHours()).padStart(2, '0') + ':' + String(v.getMinutes()).padStart(2, '0') : tx(v).slice(0, 5));
  const localDe = (s) => { const n = semAcento(s); return E.locais.find((l) => semAcento(l.nome) === n)?.nome || titulo(s); };
  const ok = [], fora = [];
  for (const r of linhas.slice(hi + 1)) {
    const nome = tx(g(r, 'nome')); if (!nome) continue;
    const cpf = tx(g(r, 'cpf')).replace(/\D/g, '').padStart(11, '0');
    const ini = data(g(r, 'ini'));
    if (!/^\d{11}$/.test(cpf) || cpf === '00000000000') { fora.push([nome, 'sem CPF']); continue; }
    if (!ini) { fora.push([nome, 'sem data de início']); continue; }
    const mat = tx(g(r, 'mat')), doc = tx(g(r, 'doc')).toUpperCase(), ex = tx(g(r, 'exame')).toUpperCase(), rem = tx(g(r, 'rem')).toUpperCase();
    const decl = /DECLIN|DESIST|N.O SEGUIR/.test(semAcento(mat + ' ' + doc + ' ' + ex)) ? (mat || doc || ex) : '';
    const mAso = ex.match(/APT[AO]\s*(?:DIA\s*)?(\d{2})\/(\d{2})(?:\/(\d{2,4}))?/);
    const pcdTx = tx(g(r, 'pcd')).toUpperCase();
    ok.push({
      nome, cpf, email: tx(g(r, 'email')), telefone: tx(g(r, 'tel')), estado_civil: /CASAD/.test(doc) ? 'Casado(a)' : 'Solteiro(a)',
      tipo_contrato: /ESTAGI/.test(semAcento(g(r, 'cargo'))) ? 'Estágio' : /APRENDIZ/.test(semAcento(g(r, 'cargo'))) ? 'Jovem aprendiz' : 'CLT',
      pcd: /^SIM/.test(semAcento(pcdTx)) ? (['Auditiva', 'Física', 'Intelectual', 'Visual', 'Mental', 'Reabilitado'].find((t) => semAcento(pcdTx).includes(semAcento(t).slice(0, 5))) || 'Física') : 'Não',
      cargo: titulo(tx(g(r, 'cargo'))) || 'Não informado', setor: titulo(tx(g(r, 'setor'))), local: localDe(tx(g(r, 'local')) || tx(g(r, 'integ'))),
      integ: titulo(tx(g(r, 'integ'))), dias_integ: tx(g(r, 'dias')).replace(/\D/g, ''), hora_integ: hora(g(r, 'hora')), data_inicio: ini,
      escala: tx(g(r, 'escala')), posicao: tx(g(r, 'pos')), processo: tx(g(r, 'proc')), requisitante: titulo(tx(g(r, 'req'))), analista_rs: titulo(tx(g(r, 'an'))),
      analista_dp: titulo(tx(g(r, 'dp'))), camisa: tx(g(r, 'camisa')), fretado: /^N\/?A$/i.test(tx(g(r, 'fret'))) ? '' : titulo(tx(g(r, 'fret'))),
      ti_itens: itensTi(semAcento(g(r, 'cham'))),
      remuneracao: /OK/.test(rem) ? 'ok' : /OCUPAD/.test(rem) ? 'ocupada' : 'pendente',
      aso_status: /APT/.test(ex) ? 'apto' : /AGEND/.test(ex) ? 'agendado' : 'aguardando',
      aso_data: mAso ? `${mAso[3] ? (mAso[3].length === 2 ? '20' + mAso[3] : mAso[3]) : new Date().getFullYear()}-${mAso[2]}-${mAso[1]}` : '',
      matricula: /^\d+$/.test(mat) ? mat : '', declinou_motivo: decl,
      docs_ok: /OK|CADASTRAD/.test(doc) || /^\d+$/.test(mat),
      observacao: [tx(g(r, 'doc')) && 'Documentação: ' + tx(g(r, 'doc')), tx(g(r, 'exame')) && 'Exame: ' + tx(g(r, 'exame'))].filter(Boolean).join(' · '),
    });
  }
  E.previa = { arquivo: file.name, ok, fora };
  modalPrevia();
}
// "NOTEBOOK, ACESSO REDE, TEAMS, E-MAIL. RH 7979" → itens da lista fechada
function itensTi(t) {
  const m = [[/NOTE?B?OOK|NOTBOOK/, 'Notebook'], [/REDE/, 'Usuário de rede'], [/E-?MAIL/, 'E-mail'], [/TEAMS/, 'Teams'], [/CELULAR/, 'Celular']];
  return m.filter(([re]) => re.test(t)).map(([, n]) => n);
}
function modalPrevia() {
  const p = E.previa; const n = (f) => p.ok.filter(f).length;
  $('#camada').innerHTML = `<div class="modal-veu"><div class="modal" style="width:min(860px,100%)" role="dialog" aria-modal="true"><h2>Importar ${esc(p.arquivo)}</h2>
   <p class="apoio">Nenhum e-mail sai na importação: quem veio da planilha não recebe boas-vindas nem gera aviso às áreas. CPF que já está na plataforma é pulado.</p>
   <div class="grade g4"><div class="caixa kpi"><b class="num">${p.ok.length}</b><span>linhas prontas</span></div><div class="caixa kpi"><b class="num">${n((x) => x.matricula)}</b><span>com matrícula</span></div><div class="caixa kpi"><b class="num">${n((x) => x.declinou_motivo)}</b><span>declínios</span></div><div class="caixa kpi"><b class="num">${p.fora.length}</b><span>ficam de fora</span></div></div>
   ${p.fora.length ? `<details><summary class="apoio">Ver as ${p.fora.length} que ficam de fora</summary><div class="tab-wrap" style="max-height:220px;overflow-y:auto"><table class="prev-import"><tbody>${p.fora.map(([nm, m]) => `<tr><td>${esc(titulo(nm))}</td><td>${esc(m)}</td></tr>`).join('')}</tbody></table></div></details>` : ''}
   <div class="tab-wrap" style="max-height:260px;overflow-y:auto"><table class="prev-import"><thead><tr><th>Nome</th><th>Início</th><th>Local</th><th>Docs</th><th>ASO</th><th>Matrícula</th></tr></thead><tbody>${p.ok.slice(0, 40).map((x) => `<tr><td>${esc(titulo(x.nome))}</td><td class="num">${dm(x.data_inicio)}</td><td>${esc(x.local)}</td><td>${x.docs_ok ? 'ok' : 'pendente'}</td><td>${esc(x.aso_status)}</td><td class="num">${esc(x.matricula || (x.declinou_motivo ? 'declinou' : '—'))}</td></tr>`).join('')}</tbody></table></div>
   <div class="linha-flex" style="justify-content:flex-end"><button class="bt" data-act="fecharModal">Cancelar</button><button class="bt pri" data-act="confirmarImport" ${p.ok.length ? '' : 'disabled'}>Importar ${plural(p.ok.length, 'linha', 'linhas')}</button></div></div></div>`;
}
async function confirmarImport(bt) {
  bt.disabled = true; bt.textContent = 'Importando…';
  let imp = 0, pul = 0; const motivos = [];
  for (let i = 0; i < E.previa.ok.length; i += 40) {
    const r = await tentar(() => api.importar(E.previa.ok.slice(i, i + 40)));
    if (!r) break; imp += r.importados; pul += r.pulados; motivos.push(...(r.motivos || []));
    bt.textContent = `Importando… ${Math.min(i + 40, E.previa.ok.length)}/${E.previa.ok.length}`;
  }
  E.previa = null; $('#camada').innerHTML = ''; E.turma = 'todas'; E.view = 'painel';
  await recarregar();
  toast(`${plural(imp, 'admissão importada', 'admissões importadas')}${pul ? ', ' + pul + ' pulada(s): ' + motivos.slice(0, 2).join('; ') : ''}`.replace(/\.?$/, '.'), pul > 0);
}

// ================================================================ EVENTOS
document.addEventListener('change', async (e) => {
  const sel = e.target.closest('[data-change="verComo"]'); if (!sel) return;
  definirVerComo(sel.value); E.view = null; await carregar(); E.view = visaoInicial(); casca();
});
document.addEventListener('click', async (e) => {
  if (!e.target.closest('.sino-wrap')) { const sl = $('#sinoLista'); if (sl && !sl.hidden) sl.hidden = true; }
  const el = e.target.closest('[data-act]'); if (!el) return;
  const a = el.dataset.act, id = el.dataset.id;
  if (a === 'fechar') { if (el.classList.contains('veu') && e.target.closest('[data-pare]')) return; $('#camada').innerHTML = ''; E.fichaAberta = null; return; }
  if (a === 'fecharModal') { $('#camada').innerHTML = ''; return; }
  switch (a) {
    case 'ir': E.view = el.dataset.v; render(); window.scrollTo(0, 0); break;
    case 'sair': await api.sair(); E.sessao = null; E.perfil = null; telaLogin(); break;
    case 'sino': abrirSino(); break;
    case 'sinoIr': { const v = el.dataset.v; $('#sinoLista').hidden = true; if (v && NAV.some((n) => n[0] === v)) { E.view = v; render(); window.scrollTo(0, 0); } if (id) abrirFicha(id); break; }
    case 'verComoSair': definirVerComo('admin'); E.view = null; await carregar(); E.view = visaoInicial(); casca(); break;
    case 'microsoft': await tentar(() => api.entrarMicrosoft()); break;
    case 'recarregar': await recarregar(); toast('Dados atualizados'); break;
    case 'tema': { const r = document.documentElement; const at = r.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); r.dataset.theme = at === 'dark' ? 'light' : 'dark'; break; }
    case 'modo': E.modo = el.dataset.m; render(); break;
    case 'ficha': abrirFicha(id); break;
    case 'importar': $('#arquivoXlsx').click(); break;
    case 'confirmarImport': await confirmarImport(el); break;
    case 'cobrar': { el.disabled = true; const r = await tentar(() => api.cobrar(id)); toast(r ? 'Cobrança na fila de e-mails' : 'Nada para cobrar: sem pendência ou sem e-mail', !r); await recarregar(); if ($('.gaveta')) abrirFicha(id); break; }
    case 'cobrarTurma': {
      el.disabled = true; let n = 0;
      for (const x of vivas().filter((y) => y.data_inicio === el.dataset.t && ['aguardando', 'pendencia'].includes(y.etapa))) { if (await tentar(() => api.cobrar(x.id))) n++; }
      toast(plural(n, 'cobrança na fila', 'cobranças na fila')); await recarregar(); break;
    }
    case 'turma': if (await tentar(() => api.atualizar(id, { data_inicio: el.dataset.t }), 'Turma alterada')) { E.turma = el.dataset.t; await recarregar(); abrirFicha(id); } break;
    case 'salvarAso': {
      const st = $('#asoSel').value;
      const p = { aso_status: st, aso_data: $('#asoData').value || null, aso_hora: ($('#asoHora')?.value || '').trim() || null, aso_local: ($('#asoLocal')?.value || '').trim() || null, aso_restricao: st === 'apto_restricao' ? ($('#asoRestr')?.value || '').trim() || null : null };
      if (st === 'agendado' && !p.aso_data) { toast('Informe a data do exame.', true); return; }
      if (st === 'apto_restricao' && !p.aso_restricao) { toast('Descreva a restrição.', true); return; }
      if (await tentar(() => api.atualizar(id, p), st === 'agendado' && adm(id)?.email ? 'ASO agendado. O candidato recebe o e-mail com data, hora e local.' : 'ASO atualizado')) { await recarregar(); if ($('.gaveta')) abrirFicha(id); else render(); }
      break;
    }
    case 'osAbrir': abrirOs(id); break;
    case 'asoRemarcar': {
      const p = { data: $('#asoData')?.value || null, hora: ($('#asoHora')?.value || '').trim() || null, local: ($('#asoLocal')?.value || '').trim() || null };
      if (!p.data && !confirm('Sem data o exame volta para "a agendar". Continuar?')) return;
      if (await tentar(() => api.asoDecidir(id, 'remarcar', p), p.data && adm(id)?.email ? 'Exame remarcado. O candidato recebe o e-mail com data, hora e local.' : 'Exame volta para a fila de agendamento')) { await recarregar(); if ($('.gaveta')) abrirFicha(id); else render(); }
      break;
    }
    case 'asoCancelar': {
      if (!confirm('Cancelar a admissão de ' + titulo(adm(id)?.nome || '') + ' por ASO inapto? As áreas recebem só o aviso de cancelamento, sem o motivo.')) return;
      if (await tentar(() => api.asoDecidir(id, 'cancelar'), 'Admissão cancelada por ASO inapto')) { await recarregar(); if ($('.gaveta')) abrirFicha(id); else render(); }
      break;
    }
    case 'tiChamado': {
      const v = ($('#tiCh-' + id)?.value || '').trim(); if (!v) { toast('Informe o número do chamado.', true); return; }
      if (await tentar(() => api.atualizar(id, { chamado_ti: v }), 'Chamado registrado')) { await recarregar(); render(); }
      break;
    }
    case 'osSalvar': case 'osValidar': {
      const p = {}; ['os_atividades', 'os_riscos', 'os_epi', 'os_epc', 'os_nrs', 'cbo'].forEach((k) => { const el2 = $('#os_' + k); if (el2) p[k] = el2.value.trim(); });
      if (a === 'osValidar' && !(p.os_atividades && p.os_riscos && p.os_epi)) { toast('Para validar, preencha ao menos atividades, riscos e EPI.', true); return; }
      if (await tentar(() => api.osSalvar(id, p, a === 'osValidar'), a === 'osValidar' ? 'Ordem de Serviço validada' : 'Ordem de Serviço salva')) { await recarregar(); $('#camada').innerHTML = ''; E.fichaAberta = null; if (E.view === 'seguranca') render(); else abrirFicha(id); }
      break;
    }
    case 'osSugerir': { const r = await tentar(() => api.osSugerir(id)); if (r === false) toast('Nenhuma linha da matriz para este cargo nesta unidade (ou a OS já está validada).', true); else if (r) { toast('Campos vazios preenchidos pela matriz'); await recarregar(); abrirOs(id); } break; }
    case 'mosSalvar': {
      const box = document.querySelector(`[data-mos="${el.dataset.i}"]`); const mo = (E.matrizOs || [])[+el.dataset.i]; const o = mo ? { id: mo.id } : {};
      box.querySelectorAll('[data-f]').forEach((x) => { o[x.dataset.f] = x.value.trim() || null; });
      if (!o.unidade || !o.cargo) { toast('Informe a unidade e o cargo.', true); return; }
      if (await tentar(() => api.salvarMatrizOs(Object.assign({ ativo: true }, o)), 'Linha da matriz salva')) await recarregar(); break;
    }
    case 'mosApagar': if (await tentar(() => api.apagarMatrizOs(+id || id), 'Linha removida')) await recarregar(); break;
    case 'remun': if (await tentar(() => api.atualizar(id, { remuneracao: el.dataset.v }), el.dataset.v === 'ok' ? 'Remuneração validada' : 'Marcada como posição ocupada')) { await recarregar(); if ($('.gaveta')) abrirFicha(id); } break;
    case 'matricular': {
      const v = ($('#matIn')?.value || '').trim();
      if (!/^\d{3,10}$/.test(v)) { toast('Digite a matrícula gerada no LG, só números.', true); return; }
      if (await tentar(() => api.atualizar(id, { matricula: v }), 'Matrícula registrada. ' + (avisadas('matricula', adm(id)) || 'Nenhuma área com e-mail para avisar.'))) { await recarregar(); abrirFicha(id); }
      break;
    }
    case 'selMail': E.selMail = +id || id; render(); break;
    case 'reenviar': await tentar(() => api.reenviar(+id), 'Volta para a fila'); await recarregar(); break;
    case 'reenviarZap': await tentar(() => api.reenviarWhatsapp(+id), 'Volta para a fila do WhatsApp'); await recarregar(); break;
    case 'togArea': { const ar = E.areas.find((x) => x.id === id); if (await tentar(() => api.salvarArea(id, { ativo: !ar.ativo }))) { ar.ativo = !ar.ativo; render(); } break; }
    case 'salvarArea': {
      const emails = $('#em_' + CSS.escape(id)).value.split(/[,;\s]+/).map((x) => x.trim().toLowerCase()).filter(Boolean);
      const ruins = emails.filter((x) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
      if (ruins.length) { toast('E-mail inválido: ' + ruins.join(', '), true); return; }
      if (await tentar(() => api.salvarArea(id, { emails }), 'E-mails salvos')) { E.areas.find((x) => x.id === id).emails = emails; render(); }
      break;
    }
    case 'entregar': if (await tentar(() => api.marcarEntrega(id, el.dataset.area, E.perfil.nome), 'Marcado como pronto')) await recarregar(); break;
    case 'fichaBusca': $('#busca').value = ''; buscar(); abrirFicha(id); break;
    case 'copiarEsteira': {
      const turma = $('#turmaKit')?.value; const texto = listaParaEsteira(turma);
      try { await navigator.clipboard.writeText(texto); toast('Lista copiada. Na Esteira, cole no passo 1 em "Colar da planilha" e clique em Importar.'); }
      catch { toast('Não consegui copiar sozinho. Abra a Esteira em tela cheia e tente de novo.', true); }
      break;
    }
    case 'enviarArq': E.alvoArq = { adm: el.dataset.adm, tipo: el.dataset.tipo }; $('#arquivoDoc').click(); break;
    case 'verArq': {
      const janela = window.open('', '_blank');   // abre já no clique, senão o navegador bloqueia
      const link = await tentar(() => api.verArquivo(el.dataset.id, el.dataset.c));
      if (link && janela) janela.location = link; else if (link) location.assign(link); else janela?.close();
      break;
    }
    case 'tirarArq':
      if (!confirm('Remover este arquivo? Ele sai do documento e do armazenamento.')) return;
      if (await tentar(() => api.removerArquivo(id, el.dataset.c), 'Arquivo removido')) { await recarregar(); abrirFicha(el.dataset.adm); }
      break;
    // ---------- piloto 19/10: dependentes (cadastro e ficha)
    case 'nDepAdd': E.depsNovo.push({}); $('#nDeps').insertAdjacentHTML('beforeend', linhaDep({}, E.depsNovo.length - 1, 'n')); break;
    case 'nDepDel': E.depsNovo.splice(+el.dataset.i, 1); $('#nDeps').innerHTML = E.depsNovo.map((d, i) => linhaDep(d, i, 'n')).join(''); $('#fNovo')?.dispatchEvent(new Event('change')); break;
    case 'fDepAdd': E.depsFicha.push({}); $('#fDeps').insertAdjacentHTML('beforeend', linhaDep({}, E.depsFicha.length - 1, 'f')); break;
    case 'fDepDel': {
      const d = E.depsFicha[+el.dataset.i];
      if (d && d.id) { if (!confirm('Remover ' + titulo(d.nome) + ' dos dependentes? Os documentos dele(a) saem da lista.')) return; if (!await tentar(() => api.apagarDependente(d.id), 'Dependente removido')) return; }
      E.depsFicha.splice(+el.dataset.i, 1); await recarregar(); abrirFicha(E.fichaAberta); break;
    }
    case 'fDepSalvar': {
      const ruins = E.depsFicha.filter((d) => (d.nome || '').trim().split(/\s+/).length < 2 || !d.parentesco || (d.cpf && !cpfValido(d.cpf)));
      if (ruins.length) { toast('Cada dependente precisa de nome completo, parentesco e, se informado, CPF válido.', true); return; }
      let n = 0;
      for (const d of E.depsFicha) { if (await tentar(() => api.salvarDependente({ id: d.id, admissao_id: id, nome: d.nome.trim(), parentesco: d.parentesco, nasc: d.nasc || null, cpf: (d.cpf || '').replace(/\D/g, '') || null, universitario: !!d.universitario, plano_saude: !!d.plano_saude, odonto: !!d.odonto, salario_familia: !!d.salario_familia }))) n++; }
      toast(plural(n, 'dependente salvo', 'dependentes salvos')); await recarregar(); abrirFicha(id); break;
    }
    // ---------- LG, eSocial, chamado, dossiê
    case 'lgMovimentado': if (await tentar(() => api.atualizar(id, { lg_movimentado_em: new Date().toISOString() }), 'Marcado como movimentado no LG')) { await recarregar(); abrirFicha(id); } break;
    case 'lgDesfazer': if (await tentar(() => api.atualizar(id, { lg_movimentado_em: null }), 'Movimentação desfeita')) { await recarregar(); abrirFicha(id); } break;
    case 'salvarRegistro': {
      const p = { esocial_enviado_em: $('#esocialData').value || null, esocial_evento: $('#esocialEv').value || null, chamado_ti: $('#chamadoIn').value.trim() || null };
      if (await tentar(() => api.atualizar(id, p), 'Registro salvo')) { await recarregar(); abrirFicha(id); } break;
    }
    case 'dossie': await exportarDossie(id, el); break;
    case 'copiarChamado': {
      const a = adm(id); const l = localDe(a.local);
      const texto = ['Solicitação de equipamentos e acessos para admissão', '',
        'Nome: ' + titulo(a.nome), 'Matrícula: ' + a.matricula, 'Cargo: ' + a.cargo + (a.setor ? ' · ' + a.setor : ''),
        'Local: ' + a.local + (l.tipo ? ' (' + l.tipo + ')' : ''), 'Centro de custo: ' + ([a.cc_codigo, a.cc_nome].filter(Boolean).join(' · ') || 'não informado'),
        'Gestor: ' + (a.requisitante || '—') + (a.requisitante_email ? ' · ' + a.requisitante_email : ''),
        'Data de início: ' + dm(a.data_inicio) + (a.formato_integ === 'online' ? ' (integração on-line)' : ' (presencial)'),
        'Itens: ' + (a.ti_itens || []).join(', ')].concat(a.substituicao ? ['Substituição de: ' + (a.pessoa_substituida || '—') + ' (verificar devolução do equipamento)'] : [])
        .concat(['', 'Aberto pela Admissão · Plataforma de Admissão']).join('\n');
      try { await navigator.clipboard.writeText(texto); toast('Texto do chamado copiado. Cole no Jira e depois registre o número aqui.'); }
      catch { await perguntar('Copie o texto abaixo', texto); }
      break;
    }
    // ---------- lista da turma
    case 'listaXlsx': try { XLSX.writeFile(planilhaLista(E.listaAtual || []), nomeLista(E.turmaLista, E.locaisLista) + '.xlsx'); } catch (e2) { toast(e2.message, true); } break;
    case 'listaCsv': { const csv = matrizLista(E.listaAtual || []).map((r) => r.map((v) => '"' + String(v).replace(/"/g, '""') + '"').join(';')).join('\r\n'); baixarBlob(nomeLista(E.turmaLista, E.locaisLista) + '.csv', new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' })); break; }
    case 'listaEnviar': {
      if (!confirm(`Enviar a lista de ${dmd(E.turmaLista)} (${E.locaisLista.join(', ')}) para a fila de e-mails?`)) return;
      el.disabled = true;
      let nome = null, b64 = null;
      try { nome = nomeLista(E.turmaLista, E.locaisLista) + '.xlsx'; b64 = XLSX.write(planilhaLista(E.listaAtual || []), { type: 'base64', bookType: 'xlsx' }); } catch (e2) { toast('A planilha não foi anexada: ' + e2.message + '. O e-mail vai só com a lista no corpo.', true); }
      const r = await tentar(() => api.enviarListaTurma(E.turmaLista, E.locaisLista, nome, b64));
      if (r) toast(`Lista na fila de e-mails: ${plural(r.pessoas, 'pessoa', 'pessoas')} para ${plural(r.destinatarios, 'destinatário', 'destinatários')}.`);
      el.disabled = false; await recarregar(); break;
    }
    case 'turmaEditar': E.turmaEdit = E.turmasCal.find((t) => t.id === id); render(); $('#cronograma')?.scrollIntoView({ block: 'start' }); break;
    case 'turmaCancelar': E.turmaEdit = null; render(); break;
    case 'turmaApagar': if (confirm('Apagar esta turma do cronograma?') && await tentar(() => api.apagarTurma(id), 'Turma apagada')) await recarregar(); break;
    // ---------- BPs e destinatários da lista
    case 'bpSalvar': {
      const box = document.querySelector(`[data-bp="${el.dataset.i}"]`); const o = E.bps[+el.dataset.i] ? { id: E.bps[+el.dataset.i].id } : {};
      box.querySelectorAll('[data-f]').forEach((x) => { o[x.dataset.f] = x.value.trim(); }); o.area = o.area || null; o.email = (o.email || '').toLowerCase();
      if (!o.nome || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(o.email)) { toast('Informe o nome e um e-mail válido.', true); return; }
      if (await tentar(() => api.salvarBp(Object.assign({ ativo: true }, o)), 'BP salvo')) await recarregar(); break;
    }
    case 'bpApagar': if (await tentar(() => api.apagarBp(+id || id), 'BP removido')) await recarregar(); break;
    case 'destSalvar': {
      const box = document.querySelector(`[data-dest="${el.dataset.i}"]`); const o = E.dest[+el.dataset.i] ? { id: E.dest[+el.dataset.i].id } : {};
      box.querySelectorAll('[data-f]').forEach((x) => { o[x.dataset.f] = x.type === 'checkbox' ? x.checked : x.value.trim(); });
      o.emails = String(o.emails || '').split(/[,;\s]+/).map((x) => x.trim().toLowerCase()).filter(Boolean);
      const ruins = o.emails.filter((x) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)); if (ruins.length) { toast('E-mail inválido: ' + ruins.join(', '), true); return; }
      if (!o.emails.length) { toast('Informe pelo menos um e-mail ou lista.', true); return; }
      if (await tentar(() => api.salvarDestinatario(Object.assign({ ativo: true }, o)), 'Destinatários salvos')) await recarregar(); break;
    }
    case 'destApagar': if (await tentar(() => api.apagarDestinatario(+id || id), 'Removido')) await recarregar(); break;
    case 'selTermo': E.selTermo = id; render(); break;
  }
});

// Envio de arquivos: prepara, guarda, pede a leitura e mostra o parecer
async function enviarArquivos(arquivos) {
  const { adm: aid, tipo } = E.alvoArq || {}; if (!aid || !arquivos.length) return;
  const nomeDoc = (t) => (E.docsFicha || []).find((d) => d.tipo === t)?.nome || t;
  const lidos = [];
  for (const original of arquivos) {
    let pronto;
    try { pronto = await prepararArquivo(original); } catch (e) { toast(e.message, true); continue; }
    toast(`Enviando ${original.name}…`);
    let idArq;
    try { idArq = await api.subirArquivo(aid, tipo, original, pronto); }
    catch (e) {
      toast(e.repetidoEm ? `${original.name}: este mesmo arquivo já está em "${nomeDoc(e.repetidoEm)}". Não entra duas vezes.` : e.message, true);
      continue;
    }
    await abrirFicha(aid);           // mostra "IA lendo…" na hora
    toast(`IA lendo ${original.name}…`);
    let r = null;
    try { r = await api.lerArquivo(idArq); }
    catch (e) {
      await api.leituraFalhou(idArq, e.message).catch(() => {});
      lidos.push(`${original.name}: a leitura automática falhou (${e.message}). Ficou para a analista conferir.`);
    }
    if (r) lidos.push(`${original.name}: ${r.motivo || r.status}`);
  }
  await recarregar();
  if (E.fichaAberta === aid) await abrirFicha(aid);
  if (lidos.length) toast(lidos.join(' | '), lidos.some((l) => !/^.*: Tudo certo/.test(l)));
}
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !document.querySelector('.modal-veu form')) $('#camada').innerHTML = ''; });
$('#arquivoDoc').addEventListener('change', async (e) => { const fs = [...e.target.files]; e.target.value = ''; await enviarArquivos(fs); });
$('#arquivoXlsx').addEventListener('change', async (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) { try { await lerPlanilha(f); } catch (err) { toast('Não consegui ler esta planilha: ' + err.message, true); } } });
setInterval(() => { if (E.perfil && !$('#camada').innerHTML && !document.querySelector('.modal-veu') && ['inicio', 'painel', 'documentos', 'entregas', 'emails', 'compliance'].includes(E.view)) recarregar(); }, 120000);

iniciar();

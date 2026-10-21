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
  // o Supabase devolve no máximo 1000 linhas por pedido: junta as páginas
  const paginas = async (q) => { let out = []; for (let i = 0; ; i += 1000) { const r = ok(await q().range(i, i + 999)); out = out.concat(r || []); if (!r || r.length < 1000) return out; } };
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
    // fluxo da Saúde (SQL 26): cadastro no SOC, contato, vacinas, modo do exame e o que vai ao R&S
    async socMarcar(adm, desfazer) { return ok(await sb.rpc('soc_marcar', { adm, desfazer: !!desfazer })); },
    async saudeAcao(adm, acao, valor, texto) { return ok(await sb.rpc('saude_acao', { adm, acao, valor: valor ?? null, texto: texto ?? null })); },
    async rsDecidirSaude(adm, decisao, texto) { return ok(await sb.rpc('rs_decidir_saude', { adm, decisao, texto: texto ?? null })); },
    async saudeParametro(chave, valor) { return ok(await sb.rpc('saude_parametro', { p_chave: chave, p_valor: valor })); },
    // agenda da Helaine (SQL 27): listas, exceção ao corte, conferências, sindicato, sino
    async listas() { return ok(await sb.from('listas').select('*').order('tipo').order('ordem').order('valor')); },
    async salvarLista(o) { return ok(await sb.from('listas').upsert(o).select('id')); },
    async apagarLista(id) { return ok(await sb.from('listas').delete().eq('id', id)); },
    async excecoes() { return ok(await sb.from('excecoes_corte').select('*').order('criado_em', { ascending: false }).limit(200)); },
    async excecaoLiberar(cpf, nome, data, motivo) { return ok(await sb.rpc('excecao_liberar', { p_cpf: cpf, p_nome: nome, p_data: data, p_motivo: motivo })); },
    async excecaoCancelar(id) { return ok(await sb.rpc('excecao_cancelar', { p_id: id })); },
    async docsRevisar(adm, desfazer) { return ok(await sb.rpc('docs_revisar', { adm, desfazer: !!desfazer })); },
    async cobrarLg(adm) { return ok(await sb.rpc('cobrar_lg', { adm })); },
    async avisarTurmaConcluida(turma, forcar) { return ok(await sb.rpc('avisar_turma_concluida', { p_turma: turma, forcar: !!forcar })); },
    async regrasSindicato() { return ok(await sb.from('regras_sindicato').select('*').order('local')); },
    async salvarRegraSindicato(o) { return ok(await sb.from('regras_sindicato').upsert(o).select('id')); },
    async apagarRegraSindicato(id) { return ok(await sb.from('regras_sindicato').delete().eq('id', id)); },
    async sindicatoConferir(adm, sindicato, obs) { return ok(await sb.rpc('sindicato_conferir', { adm, p_sindicato: sindicato, p_obs: obs || null })); },
    async comentar(adm, texto) { return ok(await sb.rpc('comentar', { adm, p_texto: texto })); },
    async notificacoes() { return ok(await sb.from('notificacoes').select('*').order('criado_em', { ascending: false }).limit(100)); },
    async notificacoesLidas(ids) { return ok(await sb.rpc('notificacoes_lidas', { ids: ids || null })); },
    // Benefícios (SQL 28)
    async beneficiosCatalogo() { return ok(await sb.from('beneficios_catalogo').select('*').eq('ativo', true).order('ordem')); },
    async beneficiosPessoa() { return paginas(() => sb.from('beneficios_pessoa').select('*').order('admissao_id').order('beneficio')); },
    async beneficioAtualizar(adms, beneficio, status, pendencia, plano, wkf) { const a = { adms, p_beneficio: beneficio, p_status: status, p_pendencia: pendencia || null, p_plano: plano || null }; if (wkf) a.p_wkf = wkf; return ok(await sb.rpc('beneficio_atualizar', a)); },
    // linhas de transporte do LG e itinerários do fretado (SQL 31)
    async transporteLinhas() { return ok(await sb.from('transporte_linhas').select('codigo_lg,nome,fornecedor,unidade,tarifa,ativo').eq('ativo', true).order('codigo_lg')); },
    async itinerarios() { return ok(await sb.from('fretado_itinerarios').select('*').eq('ativo', true).order('unidade').order('turno').order('linha')); },
    async dependentesTodos() { return paginas(() => sb.from('dependentes_visao').select('*').order('admissao_id').order('ordem').order('id')); },
    async opcoesTodas() { return paginas(() => sb.from('opcoes_beneficios').select('admissao_id,dados,recebido_em').order('admissao_id')); },
    async salvarBp(o) { return ok(await sb.from('bps').upsert(o).select('id')); },
    async apagarBp(id) { return ok(await sb.from('bps').delete().eq('id', id)); },
    async areasDestino() { return ok(await sb.from('areas_destino').select('*').order('area_id').order('unidade')); },
    async salvarAreaDestino(o) { return ok(await sb.from('areas_destino').upsert(o).select('id')); },
    async apagarAreaDestino(id) { return ok(await sb.from('areas_destino').delete().eq('id', id)); },
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
  if (/listas_unica|regras_sindicato_unica/.test(m || '')) return 'Este item já existe na lista.';
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
const TIPO_EMAIL = { cadastro: 'Novo aprovado', boasvindas: 'Boas-vindas', cobranca: 'Cobrança', docs_ok: 'Documentação completa', matricula: 'Matrícula', lg: 'Movimentar no LG', convite: 'Convite da integração', turma: 'Lista da turma', beneficios_opcoes: 'Opções de benefícios', termo: 'Termo de Transparência', via_assinada: 'Via assinada do kit', aso: 'Exame admissional (ASO)', aso_inapto: 'ASO inapto', lg_cobranca: 'Cobrança: movimentar no LG', itinerario: 'Itinerário do fretado', turma_concluida: 'Admissões concluídas', exame_etapa: 'Etapa de exame médico', saude_rs: 'Saúde → R&S', declinou: 'Admissão cancelada' };
// a jornada de e-mails do colaborador, na ordem (os demais e-mails vão para as áreas e a equipe)
const JORNADA_EMAILS = [
  ['boasvindas', 'Cadastro do aprovado pelo R&S', 'Boas-vindas: data da integração, lista de documentos, como enviar (LG), prazo do corte da turma, link do Código de Conduta e aviso de privacidade.'],
  ['cobranca', 'Todo dia, enquanto faltar documento (até o corte da turma)', 'Cobrança dos documentos que faltam ou precisam ser reenviados; também pelo botão "Cobrar" da ficha.'],
  ['exame_etapa', 'Cadastro no SOC marcado', 'Etapa de exame médico (card da Saúde): contato pelo WhatsApp, caderneta de vacinação com as vacinas do parâmetro vacinas_obrigatorias, não faltar, clínica parceira ou reembolso.'],
  ['itinerario', 'Cadastro com linha de fretado (ou fretado na integração)', 'Itinerário do fretado: o da integração (administrativo) e o do turno, com o horário de cada ponto e o link do buscador do ponto mais perto (CEP, rua ou bairro).'],
  ['convite', 'Matrícula registrada', 'Convite da integração: data, hora, local (ou link do Teams) e quem procurar ao chegar; gestor e recrutadora em cópia.'],
  ['via_assinada', 'Colaborador assina o kit e escolhe e-mail', 'Via assinada do kit admissional em PDF, com o código de verificação.'],
];
// e-mails que dependem de um acontecimento
const EMAILS_EVENTO = [
  ['aso', 'Saúde agenda o exame (ASO agendado com data)', 'Ao candidato: data, hora, local e orientações do exame (parâmetro aso_orientacoes). Se o ASO fica inapto, a analista de R&S e a Admissão recebem o aviso.'],
  ['saude_rs', 'Saúde envia um caso ao R&S', 'À analista de R&S (sem ela, à lista do R&S): sem contato há 24h, falta ao exame, limite de remarcação, recusa de vacina ou erro no cadastro do SOC. O R&S decide na ficha.'],
  ['declinou', 'Admissão cancelada', 'Às áreas com aviso no gatilho "admissão cancelada" (TI: cancelar chamado, acessos e equipamento).'],
];
const EMAILS_EQUIPE = [
  ['cadastro', 'Cadastro', 'Áreas com aviso no cadastro (Saúde para agendar o ASO, Admissão, Remuneração, Sindical no Comercial, TI previsão de equipamento).'],
  ['lg', 'ASO apto', 'Pedido à analista de R&S para movimentar a pessoa no LG.'],
  ['lg_cobranca', 'Admissão clica em "Cobrar a recrutadora"', 'Cobrança à analista de R&S responsável: a pessoa está apta e ainda não foi movimentada no LG.'],
  ['turma_concluida', 'Última matrícula da turma registrada', 'Às áreas: admissões da turma concluídas, com nome, cargo, local, matrícula e centro de custo (uma vez por turma; a Admissão pode reenviar).'],
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
  tLinhas: [], itins: null, itinUni: 'UNIDADE 1', itinTurno: '1º TURNO',
  listas: [], excecoes: [], notifs: [], regrasSind: [], benCat: [], benPessoa: [], depsTodos: [], opcTodas: [], benTurma: null, listaTipo: 'analista_rs', iniF: { setor: '', local: '', mes: '' }, sindTodos: false,
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
  ['esp-sindical', 'Especialista · Relações Sindicais', ['especialista'], 'sindical'],
  ['esp-beneficios', 'Especialista · Benefícios', ['especialista'], 'beneficios'],
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
  const esp = so && tem('especialista') ? minhaArea() : '';
  return so && tem('saude') ? 'saude' : so && tem('remuneracao') ? 'remuneracao' : so && tem('signatario') ? 'kit' : so && tem('compliance') ? 'compliance'
    : esp === 'sindical' ? 'sindicatos' : ['beneficios', 'beneficios_prev', 'beneficios_cancela'].includes(esp) ? 'beneficios' : navVisivel()[0]?.[0];
}

// [id, nome, ícone, quem vê, seção do menu]
const NAV = [
  ['inicio', 'Início', 'inicio', () => tem('rs', 'admissao'), 1],
  ['novo', 'Cadastrar aprovado', 'novo', () => tem('rs', 'admissao'), 1],
  ['painel', 'Painel da admissão', 'painel', () => tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario'), 1],
  ['turma', 'Turma e lista', 'esp', () => tem('rs', 'admissao'), 1],
  ['cronograma', 'Cronograma e corte', 'esp', () => tem('rs', 'admissao'), 1],
  ['documentos', 'Documentos', 'docs', () => tem('admissao'), 1],
  ['compliance', 'Termo de Transparência', 'docs', () => tem('compliance'), 1],
  ['kit', 'Esteira de assinatura', 'kit', () => tem('admissao', 'signatario'), 1],
  ['remuneracao', 'Remuneração', 'remun', () => tem('remuneracao', 'admissao'), 1],
  ['saude', 'Saúde ocupacional', 'saude', () => tem('saude'), 1],
  ['beneficios', 'Benefícios', 'esp', () => souBeneficios() && temSql28(), 1],
  ['sindicatos', 'Sindicatos', 'esp', () => souSindical() && temSql27(), 1],
  ['seguranca', 'Segurança do Trabalho', 'esp', () => souSeguranca(), 1],
  ['ti', 'TI · equipamentos e acessos', 'esp', () => souTi(), 1],
  ['entregas', 'Minhas entregas', 'esp', () => tem('especialista', 'admissao'), 1],
  ['especialistas', 'Especialistas e avisos', 'esp', () => tem('admissao'), 2],
  ['listas', 'Listas e regras', 'esp', () => tem('admissao') && temSql27(), 2],
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
async function aposSaude(id) { await recarregar(); if ($('.gaveta')) abrirFicha(id); else render(); }
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
  if (tem('rs', 'admissao', 'saude', 'remuneracao', 'signatario') || souSeguranca() || souTi() || souSindical() || souBeneficios()) {
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
    if (api.areasDestino) tarefas.push(api.areasDestino().then((r) => (E.areaDest = r || [])).catch(() => (E.areaDest = null)));
    tarefas.push(api.destinatariosTurma().then((r) => (E.dest = r || [])));
  }
  if (tem('compliance')) tarefas.push(api.termos().then((r) => (E.termos = r || [])));
  // SQL 27 e 28: cada carga falha calada se o SQL ainda não foi rodado
  const calado = (f, k) => { if (f) tarefas.push(f().then((r) => (E[k] = r || [])).catch(() => (E[k] = []))); };
  calado(api.listas && (() => api.listas()), 'listas');
  calado(api.notificacoes && (() => api.notificacoes()), 'notifs');
  calado(api.regrasSindicato && (() => api.regrasSindicato()), 'regrasSind');
  calado(api.transporteLinhas && (() => api.transporteLinhas()), 'tLinhas');
  if (tem('rs', 'admissao')) calado(api.excecoes && (() => api.excecoes()), 'excecoes');
  if (souBeneficios()) {
    calado(api.beneficiosCatalogo && (() => api.beneficiosCatalogo()), 'benCat');
    calado(api.beneficiosPessoa && (() => api.beneficiosPessoa()), 'benPessoa');
    calado(api.dependentesTodos && (() => api.dependentesTodos()), 'depsTodos');
    calado(api.opcoesTodas && (() => api.opcoesTodas()), 'opcTodas');
  }
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
  (E.notifs || []).filter((n) => Date.now() - new Date(n.criado_em).getTime() < 14 * 864e5)
    .forEach((n) => l.push({ chave: 'db:' + n.id, nivel: n.origem === 'mencao' ? 0 : 2, texto: n.texto, id: n.admissao_id || '', dbId: n.lida_em ? null : n.id }));
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
  if (tem('rs', 'admissao', 'saude', 'remuneracao') || souSeguranca() || souTi() || souSindical() || souBeneficios()) listaAtencao().forEach(([p, a, tx]) => l.push({ chave: 'at:' + a.id + ':' + tx.slice(0, 24), nivel: p, texto: nome(a.nome) + ' · ' + tx, id: a.id }));
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
  const lidas = l.filter((x) => x.dbId).map((x) => x.dbId);
  if (lidas.length && api.notificacoesLidas) { (E.notifs || []).forEach((n) => { if (lidas.includes(n.id)) n.lida_em = new Date().toISOString(); }); api.notificacoesLidas(lidas).catch(() => {}); }
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
      if (api.notificacoes && temSql27()) tarefas.push(api.notificacoes().then((r) => (E.notifs = r || [])).catch(() => {}));
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
// fluxo da Saúde (SQL 26, documentos da Saúde de 20/10/2026)
const MOTIVO_SAUDE = { sem_contato: 'sem contato há 24h', falta: 'faltou ao exame', remarcacao_excedida: 'limite de remarcação atingido', recusa_vacina: 'recusou a vacinação', erro_soc: 'erro no cadastro do SOC' };
const MODO_ASO = { credenciada: 'Rede credenciada', proxima: 'Clínica próxima ao candidato', reembolso: 'Reembolso (clínica conveniente)' };
const examesCompl = (a) => a.exames_complementares ?? a.exames_complementares_tabela ?? null;
const remarcMax = (a) => Number(E.params?.aso_remarcacoes_max || 1) + (a.aso_remarcacoes_extra || 0);
const contatoAtrasado = (a) => !!(a.soc_cadastrado_em && !a.saude_contato_em && a.saude_prazo_contato && Date.now() > new Date(a.saude_prazo_contato).getTime());
const janelaContatoVencida = (a) => !!(a.saude_sem_contato_desde && !a.saude_escalado_em && Date.now() - new Date(a.saude_sem_contato_desde).getTime() > Number(E.params?.saude_janela_contato_horas || 24) * 3600e3);
const prazoAsoPassa = (a) => !!(a.soc_cadastrado_em && a.saude_prazo_aso && !asoOk(a) && a.saude_prazo_aso > a.data_inicio);
// sem o SQL 26 rodado a visão não traz os campos: a tela fica como antes
const fluxoSaude = (a) => !!a && 'soc_cadastrado_em' in a;
const pendenteSaude = (a) => !a.declinou_em && !a.matricula && !asoOk(a);
function chipSoc(a) { if (!fluxoSaude(a) || (asoOk(a) && !a.soc_cadastrado_em)) return '—'; return a.soc_cadastrado_em ? chip('c-ok', 'ok', 'SOC ' + dm(new Date(a.soc_cadastrado_em))) : chip('c-at', 'at', 'Falta cadastrar no SOC'); }
function chipVacina(a) { if (!a.vacina_status || (asoOk(a) && a.vacina_status === 'pendente')) return ''; return a.vacina_status === 'comprovada' ? chip('c-ok', 'ok', 'Vacinas comprovadas') : a.vacina_status === 'recusada' ? chip('c-cr', 'cr', 'Recusou a vacina') : chip('c-at', 'at', 'Vacinas pendentes'); }
function chipContato(a) {
  if (!fluxoSaude(a)) return '—';
  if (!a.soc_cadastrado_em) return '<span class="apoio">depois do SOC</span>';
  if (a.saude_escalado_em) return chip('c-cr', 'cr', 'No R&S: ' + (MOTIVO_SAUDE[a.saude_escalado_motivo] || ''));
  if (a.aso_faltou_em && a.saude_sem_contato_desde) return chip(janelaContatoVencida(a) ? 'c-cr' : 'c-at', janelaContatoVencida(a) ? 'cr' : 'at', 'Faltou ' + dm(new Date(a.aso_faltou_em)) + (janelaContatoVencida(a) ? ' · 24h passaram' : ''));
  if (a.saude_sem_contato_desde) return chip(janelaContatoVencida(a) ? 'c-cr' : 'c-at', janelaContatoVencida(a) ? 'cr' : 'at', 'Não atende desde ' + quando(a.saude_sem_contato_desde));
  if (a.saude_contato_em) return chip('c-ok', 'ok', 'Contato ' + dm(new Date(a.saude_contato_em)));
  return chip(contatoAtrasado(a) ? 'c-cr' : 'c-at', contatoAtrasado(a) ? 'cr' : 'at', 'Contato até ' + dmh(a.saude_prazo_contato));
}
function chipPrazoAso(a) {
  if (!fluxoSaude(a)) return '—';
  const c = examesCompl(a); const d = c ? 10 : 5;
  const rot = (c == null ? 'função fora da tabela · ' : '') + d + ' dias úteis';
  if (!a.soc_cadastrado_em || !a.saude_prazo_aso) return `<span class="apoio" style="font-size:12px">${rot}</span>`;
  return (prazoAsoPassa(a) ? chip('c-cr', 'cr', 'até ' + dm(a.saude_prazo_aso) + ': passa da integração') : '<span class="num">até ' + dm(a.saude_prazo_aso) + '</span>') + `<br><span class="apoio" style="font-size:12px">${rot}</span>`;
}
// Segurança do Trabalho: a Ordem de Serviço vale para todas as unidades (Josianny, 20/10/2026)
const temOs = (a) => a.os_status !== 'nao_aplica';
const minhaArea = () => (E.verComo ? (E.verComo.area_id || '') : (E.perfil && E.perfil.area_id) || '');
const souSeguranca = () => tem('admissao') || (tem('especialista') && ['seguranca', 'seguranca_prev'].includes(minhaArea()));
const souTi = () => tem('admissao') || (tem('especialista') && ['ti', 'tiprev', 'ti_cancela'].includes(minhaArea()));
const souSindical = () => tem('admissao') || (tem('especialista') && minhaArea() === 'sindical');
const souBeneficios = () => tem('admissao') || (tem('especialista') && ['beneficios', 'beneficios_prev', 'beneficios_cancela'].includes(minhaArea()));
// o que já está instalado no banco (cada SQL grava um parâmetro): sem ele, a tela fica como antes
const temSql27 = () => E.params?.cobranca_imediata != null;
const temSql28 = () => (E.benCat || []).length > 0;
const temSql31 = () => (E.tLinhas || []).length > 0;
const temSql35 = () => (E.areas || []).some((a) => 'reavisar' in a);
// tipo de uso do carro e prazo das políticas de frota (Frotas, 21/10/2026)
const FROTA_TIPO = { comercial: 'Ferramenta de trabalho (comercial, até coordenador)', gerente_comercial: 'Gerente comercial (carro benefício, cartão como ferramenta)', beneficio: 'Só benefício (corporativo)' };
const frotaPrazo = (tipo) => (tipo === 'beneficio' ? 30 : 2);
// fretado (SQL 31): unidade dos itinerários por local; linhas "24 - FRETADO" do LG
const UNID_FRETADO = { 'Pouso Alegre': 'UNIDADE 1', 'Fernão Dias': 'UNIDADE 2', 'Gráfica SSBV': 'SSBV' };
const INTEG_FRETADO = ['Pouso Alegre', 'Fernão Dias'];   // fretado na integração: turno administrativo
function linhasFretado(local) {
  const u = UNID_FRETADO[local];
  return (E.tLinhas || []).filter((l) => /^24\b/.test(l.fornecedor || '') && ((u && l.unidade === u) || (local === 'Pouso Alegre' && l.codigo_lg === 71) || (local === 'CD SC' && /CHAPEC/i.test(l.nome)) || (local === 'CD BA' && /BAHIA/i.test(l.nome))));
}
const nomeLinha = (l) => String(l.nome || '').replace(/^(UNIDADE [12]|SSBV)\s*-\s*/i, '').replace(/^FRETADO\s+/i, '');
// cadastro: linhas do fretado do turno escolhido (itinerários) e as do administrativo para a integração
function renderFretado() {
  const wrap = $('#n_fretadoWrap'); if (!wrap) return;
  const local = $('#n_local').value; const u = UNID_FRETADO[local];
  const antes = $('#n_fretado')?.value || ''; const antesInteg = $('#n_frIntegSel')?.value || '';
  const pessoa = { local, turno: valLista('n_turno'), escala: valLista('n_escala'), tipo_contrato: $('#n_tipo').value };
  const doUni = u && E.itins ? E.itins.filter((i) => i.unidade === u) : [];
  const op = (i, sel) => `<option value="${esc(i.chave)}" ${i.chave === sel ? 'selected' : ''}>${esc(titulo(i.linha))} · chega ${esc(i.chegada || '—')} · ${plural((i.pontos || []).length, 'parada', 'paradas')}${i.codigo_lg ? ' · LG ' + i.codigo_lg : ''}</option>`;
  if (temSql31() && doUni.length) {
    const ts = turnosItin(pessoa);
    if (!ts.length) { wrap.innerHTML = '<select id="n_fretado" disabled><option value="">Escolha o turno (ou a escala 2x2x3) primeiro</option></select>'; $('#n_fretadoAj').textContent = 'As linhas do fretado dependem do turno.'; }
    else {
      let l = doUni.filter((i) => ts.includes(i.turno));
      const turma = (String(pessoa.escala || '').match(/2X2X3\s*-?\s*([A-D])\b/i) || [])[1];
      if (turma) { const lt = l.filter((i) => new RegExp('^TURMA ' + turma.toUpperCase() + '\\b').test(i.linha)); if (lt.length) l = lt; }
      const grupos = [...new Set(l.map((i) => i.turno))];
      wrap.innerHTML = `<select id="n_fretado" data-lista="fretado"><option value="">Não usa</option>${grupos.map((g) => (grupos.length > 1 ? `<optgroup label="${esc(g)}">` : '') + l.filter((i) => i.turno === g).map((i) => op(i, antes)).join('') + (grupos.length > 1 ? '</optgroup>' : '')).join('')}<option value="__outra">Outra (digitar)…</option></select><input id="n_fretado_outra" hidden placeholder="Linha ou bairro" style="margin-top:6px">`;
      $('#n_fretadoAj').textContent = l.length ? `Linhas do ${ts.join(' / ').toLowerCase()} em ${local}. O colaborador recebe o itinerário por e-mail.` : 'Nenhuma linha de fretado neste turno.';
    }
  } else if (temSql31() && !linhasFretado(local).length) {
    wrap.innerHTML = '<select id="n_fretado" disabled><option value="">Esta unidade não tem fretado</option></select>';
    if ($('#n_fretadoAj')) $('#n_fretadoAj').textContent = 'Fretado só nas fábricas (Pouso Alegre, Fernão Dias, SSBV) e nas linhas do LG dos CDs de SC e BA.';
  } else {
    wrap.innerHTML = selFretado(local);
    if ($('#n_fretadoAj')) $('#n_fretadoAj').textContent = temSql31() ? (linhasFretado(local).length ? 'Linhas do LG para esta unidade (sem itinerário cadastrado).' : 'Esta unidade não tem fretado.') : 'Benefícios cadastra no Rota para o QR Code.';
  }
  const box = $('#n_frIntegBox');
  if (box) {
    const adm = doUni.filter((i) => i.turno === 'ADMINISTRATIVO');
    box.hidden = !INTEG_FRETADO.includes(local) || !adm.length;
    const w = $('#n_frIntegWrap'); w.hidden = box.hidden || !$('#n_frInteg').checked;
    w.innerHTML = `<select id="n_frIntegSel" style="margin-top:6px"><option value="">Escolha a linha do administrativo</option>${adm.map((i) => op(i, antesInteg)).join('')}</select>`;
  }
  mostrarOutras($('#fNovo'));
}
function selFretado(local, cod, txt) {
  const ls = linhasFretado(local);
  if (!ls.length) return selLista('n_fretado', 'fretado', txt || '', { filtro: local, vazio: 'Não usa', outra: true, ph: 'Linha ou bairro (vazio = não usa)' });
  return `<select id="n_fretado" data-lista="fretado"><option value="">Não usa</option>${ls.map((l) => `<option value="${l.codigo_lg}" ${+cod === l.codigo_lg ? 'selected' : ''}>${esc(nomeLinha(l))} · LG ${l.codigo_lg}</option>`).join('')}<option value="__outra">Outra (digitar)…</option></select><input id="n_fretado_outra" hidden placeholder="Linha ou bairro" style="margin-top:6px">`;
}
// turno da admissão → turno(s) dos itinerários
function turnosItin(a) {
  if (a.tipo_contrato === 'Jovem aprendiz') return ['APRENDIZ'];
  const e = String(a.escala || '').toUpperCase(); const tu = String(a.turno || '').toLowerCase();
  if (/2X2X3/.test(e)) return ['2X2X3 - 1º TURNO', '2X2X3 - 2º TURNO'];
  if (tu.startsWith('1')) return ['1º TURNO'];
  if (tu.startsWith('2')) return ['2º TURNO'];
  if (tu.startsWith('3')) return ['3º TURNO SEG-SAB', '3º TURNO DOM'];
  if (tu.startsWith('adm')) return ['ADMINISTRATIVO'];
  return [];
}
async function garantirItins() {
  if (E.itins) return;
  if (!api.itinerarios || !temSql31()) { E.itins = []; return; }
  if (!E.itinsP) E.itinsP = api.itinerarios().then((r) => { E.itins = r || []; }).catch(() => { E.itins = []; }).finally(() => { E.itinsP = null; });
  await E.itinsP;
}
function itinerariosDe(a) {
  if (!E.itins) return [];
  if (a.fretado_itin_chave) return E.itins.filter((i) => i.chave === a.fretado_itin_chave);
  const u = UNID_FRETADO[a.local]; if (!u) return [];
  const ts = turnosItin(a); const doTurno = E.itins.filter((i) => i.unidade === u && (!ts.length || ts.includes(i.turno)));
  if (ts.includes('APRENDIZ')) return doTurno;
  const turma = (String(a.escala || '').match(/2X2X3\s*-?\s*([A-D])\b/i) || [])[1];
  if (turma) { const l = doTurno.filter((i) => new RegExp('^TURMA ' + turma.toUpperCase() + '\\b').test(i.linha)); if (l.length) return l; }
  const alvo = semAcento(String(a.fretado || '').replace(/^(UNIDADE [12]|SSBV)\s*-\s*/i, ''));
  const l = doTurno.filter((i) => (a.fretado_codigo && i.codigo_lg === a.fretado_codigo) || (alvo && (semAcento(i.linha).includes(alvo) || alvo.includes(semAcento(i.linha)))));
  return l.length ? l : [];
}
const itinIntegDe = (a) => (a.fretado_integ && a.fretado_integ_chave && E.itins ? E.itins.filter((i) => i.chave === a.fretado_integ_chave) : []);
function tabelaItin(i) {
  const ps = Array.isArray(i.pontos) ? i.pontos : [];
  return `<details class="itin"><summary><b>${esc(i.linha)}</b> · ${esc(i.turno)} · ${plural(ps.length, 'parada', 'paradas')} · ${ps[0] ? 'sai ' + esc(ps[0].horario || '') + ' · ' : ''}chega ${esc(i.chegada || '—')}${i.codigo_lg ? ' · LG ' + i.codigo_lg : ' · sem código no LG'}</summary>
    <div class="tab-wrap"><table><thead><tr><th>Horário</th><th>Ponto</th><th>Bairro</th><th>Endereço</th></tr></thead><tbody>${ps.map((x) => `<tr><td class="num">${esc(x.horario || '')}</td><td>${esc(x.referencia || '')}</td><td>${esc(x.bairro || '')}</td><td style="font-size:12px">${esc(String(x.endereco || '').replace(/, Brasil$/, ''))}</td></tr>`).join('')}</tbody></table></div></details>`;
}
function blocoItinerario(a) {
  if (!temSql31() || !UNID_FRETADO[a.local]) return '';
  if (!a.fretado && !a.fretado_codigo && !a.fretado_itin_chave && !a.fretado_integ) return '<span class="apoio" style="font-size:12.5px">Sem fretado informado no cadastro.</span>';
  if (!E.itins) return '<span class="apoio" style="font-size:12.5px">Carregando itinerários…</span>';
  const l = itinerariosDe(a); const ig = itinIntegDe(a);
  if (l.length || ig.length) return (ig.length ? '<b style="font-size:13px">Na integração</b>' + ig.map(tabelaItin).join('') : '') + (l.length ? '<b style="font-size:13px">No dia a dia</b>' + l.map(tabelaItin).join('') : '');
  const ts = turnosItin(a);
  return `<span class="apoio" style="font-size:12.5px">Sem itinerário para a linha ${esc(a.fretado || 'LG ' + a.fretado_codigo)}${ts.length ? ' no ' + esc(ts.join(' / ')) : ' (turno não informado)'}. Confira a linha e o turno, ou a tabela de linhas em Benefícios.</span>`;
}
// listas suspensas mantidas pela Admissão (SQL 27); filtro = unidade (linhas do fretado)
const opcoesLista = (tipo, filtro) => (E.listas || []).filter((x) => x.tipo === tipo && x.ativo !== false && (filtro == null || !x.filtro || x.filtro === filtro)).map((x) => x.valor);
// select da lista, com "Outra (digitar)"; sem lista cadastrada vira campo de texto
function selLista(id, tipo, sel, o = {}) {
  const vals = opcoesLista(tipo, o.filtro).length ? opcoesLista(tipo, o.filtro) : (o.padrao || []);
  if (!vals.length) return `<input id="${id}" value="${esc(sel || '')}" placeholder="${esc(o.ph || '')}"${o.attrs || ''}>`;
  const fora = sel && !vals.includes(sel);
  return `<select id="${id}" data-lista="${tipo}"${o.attrs || ''}>${o.vazio != null ? `<option value="">${esc(o.vazio)}</option>` : ''}${vals.map((v) => `<option ${v === sel ? 'selected' : ''}>${esc(v)}</option>`).join('')}${fora ? `<option selected>${esc(sel)}</option>` : ''}${o.outra ? '<option value="__outra">Outra (digitar)…</option>' : ''}</select>${o.outra ? `<input id="${id}_outra" hidden placeholder="${esc(o.ph || 'Digite')}" style="margin-top:6px"${o.attrs || ''}>` : ''}`;
}
const valLista = (id) => { const s = $('#' + id); if (!s) return ''; return s.value === '__outra' ? ($('#' + id + '_outra')?.value || '').trim() : String(s.value || '').trim(); };
const mostrarOutras = (raiz = document) => raiz.querySelectorAll('select[data-lista]').forEach((s) => { const i = $('#' + s.id + '_outra'); if (i) i.hidden = s.value !== '__outra'; });
// regra de sindicato do local (mesma do banco: a do cargo ganha da geral)
function sindSugerido(local, cargo) {
  const c = semAcento(cargo || '');
  const rs = (E.regrasSind || []).filter((r) => r.ativo !== false && r.local === local && (!r.cargo_contem || c.includes(semAcento(r.cargo_contem))));
  rs.sort((x, y) => (x.cargo_contem ? 0 : 1) - (y.cargo_contem ? 0 : 1) || String(y.cargo_contem || '').length - String(x.cargo_contem || '').length);
  return rs[0]?.sindicato || '';
}
const SIND_ST = { pendente: ['c-ag', 'ag', 'Sindicato não informado'], conforme: ['c-ok', 'ok', 'Conforme a regra'], divergente: ['c-at', 'at', 'Diferente da regra'], sem_regra: ['c-at', 'at', 'Sem regra para o local'], conferido: ['c-ok', 'ok', 'Conferido'] };
const chipSind = (a) => { const s = SIND_ST[a.sindicato_status] || SIND_ST.pendente; return chip(s[0], s[1], s[2]); };
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
    // fluxo da Saúde (SQL 26)
    if (fluxoSaude(a) && pendenteSaude(a)) {
      if (tem('rs', 'admissao') && !a.soc_cadastrado_em && a.aso_status === 'aguardando') atencao.push([1, a, 'Falta cadastrar no SOC: sem isso a Saúde não começa o contato']);
      if (tem('saude', 'admissao') && contatoAtrasado(a)) atencao.push([0, a, 'Contato da Saúde atrasado: o prazo era ' + dmh(a.saude_prazo_contato) + ' (regra das 12h)']);
      if (tem('saude', 'admissao') && janelaContatoVencida(a)) atencao.push([0, a, (a.aso_faltou_em ? 'Faltou ao exame' : 'Sem contato') + ' há mais de ' + (E.params?.saude_janela_contato_horas || '24') + 'h: enviar ao R&S']);
      if (tem('rs', 'admissao') && a.saude_escalado_em) atencao.push([0, a, 'A Saúde enviou ao R&S (' + (MOTIVO_SAUDE[a.saude_escalado_motivo] || '') + '): decida na ficha']);
      if (tem('saude', 'admissao') && a.vacina_status === 'pendente' && a.aso_status === 'agendado') atencao.push([1, a, 'Exame agendado e as vacinas ainda não comprovadas']);
      if (tem('saude', 'admissao', 'rs') && prazoAsoPassa(a)) atencao.push([1, a, 'Prazo do ASO (' + (examesCompl(a) ? 10 : 5) + ' dias úteis) termina depois da integração']);
    }
    if (tem('admissao') && chamadoTiAtrasado(a)) atencao.push([1, a, 'Chamado de TI ainda não aberto: o prazo da TI é até ' + dm(prazoChamadoTi(a))]);
    if (souSeguranca() && temOs(a) && a.os_status !== 'validada' && !a.declinou_em && a.etapa !== 'pronto') atencao.push([2, a, 'Ordem de Serviço da Segurança ainda não validada']);
    if (faltaLg(a)) atencao.push([1, a, 'ASO apto: falta a recrutadora movimentar no LG']);
    if (tem('admissao') && faltaDossie(a)) atencao.push([2, a, 'Kit assinado: exportar o dossiê antes do descarte']);
    if (resumo(a).analise) atencao.push([2, a, plural(resumo(a).analise, 'documento recebido para conferir', 'documentos recebidos para conferir')]);
    if ('docs_revisados_em' in a && !a.matricula && !a.docs_revisados_em && ['docs_ok', 'registro'].includes(a.etapa) && tem('admissao') && a.conferente_docs && semAcento(a.conferente_docs) === semAcento(E.perfil?.nome)) atencao.push([1, a, 'Revisão final dos documentos é sua: confira e dê o ok']);
    if ('docs_revisados_em' in a && !a.matricula && ['docs_ok', 'registro'].includes(a.etapa) && tem('admissao') && !a.conferente_docs) atencao.push([2, a, 'Documentação completa sem responsável pela conferência final']);
    if (souBeneficios() && a.era_dependente_cimed && a.matricula && !a.declinou_em && (E.benPessoa || []).some((x) => x.admissao_id === a.id && ['saude_unimed', 'saude_sulamerica'].includes(x.beneficio) && !['enviado', 'ativo', 'cancelado', 'recusado'].includes(x.status))) atencao.push([1, a, 'Benefícios: era dependente de alguém da Cimed no plano, mover para titular no WKF']);
    if (souBeneficios() && (E.benPessoa || []).length && a.matricula && a.data_inicio <= iso(new Date())) {
      const bi = (E.benPessoa || []).filter((x) => x.admissao_id === a.id);
      if (bi.some((x) => ['saude_unimed', 'saude_sulamerica'].includes(x.beneficio) && ['escolhido', 'pronto'].includes(x.status))) atencao.push([0, a, 'Benefícios: inclusão no plano de saúde a enviar (D+1 da integração)']);
      if (bi.some((x) => x.beneficio === 'fretado' && ['escolhido', 'pronto', 'enviado'].includes(x.status))) atencao.push([1, a, 'Benefícios: fretado sem QR Code confirmado no Rota']);
    }
    if ('sindicato_status' in a && !a.matricula && ['divergente', 'sem_regra'].includes(a.sindicato_status) && souSindical()) atencao.push([1, a, 'Sindicato para conferir: ' + (a.sindicato || '—') + (a.sindicato_sugerido ? ' (regra do local: ' + a.sindicato_sugerido + ')' : ' (sem regra para o local)')]);
  });
  atencao.sort((x, y) => x[0] - y[0]);
  return atencao;
}
VIEWS.inicio = () => {
  const F = E.iniF; const todos = E.admissoes;
  const setores = [...new Set(todos.map((a) => (a.setor || '').trim()).filter(Boolean))].sort();
  const locais = [...new Set(todos.map((a) => a.local).filter(Boolean))].sort();
  const meses = [...new Set(todos.map((a) => (a.data_inicio || '').slice(0, 7)).filter(Boolean))].sort();
  const noRecorte = (a) => (!F.setor || (a.setor || '').trim() === F.setor) && (!F.local || a.local === F.local);
  const base = todos.filter((a) => noRecorte(a) && (!F.mes || (a.data_inicio || '').startsWith(F.mes)));
  const n = (f, l = base) => l.filter(f).length;
  const pct = (x, d) => (d ? Math.round((x * 100) / d) : 0);
  const cand = base.length; const decl = n((a) => a.declinou_em); const admit = n((a) => !a.declinou_em && a.matricula);
  const proc = n((a) => !a.declinou_em && !a.matricula); const aptos = n((a) => !a.declinou_em && asoOk(a));
  const pendEx = n((a) => !a.declinou_em && !a.matricula && !asoOk(a)); const remOk = n((a) => a.remuneracao === 'ok');
  const ocup = n((a) => !a.declinou_em && a.remuneracao === 'ocupada'); const semRem = n((a) => !a.declinou_em && a.remuneracao === 'pendente');
  const lim = { desist: Number(param('alerta_desistencia', '15')) || 15, exames: Number(param('alerta_exames_pendentes', '60')) || 60 };
  const kpi = (v, rot, sub, alerta) => `<div class="caixa kpi${alerta ? ' alerta' : ''}"><b class="num">${v}</b><span>${rot}</span>${sub ? `<small class="apoio" style="font-size:12px">${sub}</small>` : ''}</div>`;
  const grupoSetor = (l) => Object.entries(l.reduce((m, a) => { const s = (a.setor || 'Sem setor').trim(); (m[s] = m[s] || []).push(a); return m; }, {}))
    .map(([s, x]) => ({ s, cand: x.length, adm: n((a) => !a.declinou_em && a.matricula, x), aptos: n((a) => !a.declinou_em && asoOk(a), x), pend: n((a) => !a.declinou_em && !a.matricula && !asoOk(a), x), decl: n((a) => a.declinou_em, x), rem: n((a) => a.remuneracao === 'ok', x) }))
    .sort((x, y) => y.cand - x.cand);
  const porSetor = grupoSetor(base);
  let acum = 0;
  const porMes = meses.filter((m) => !F.mes || m === F.mes).map((m) => { const x = todos.filter((a) => (a.data_inicio || '').startsWith(m) && noRecorte(a)); const adm = n((a) => !a.declinou_em && a.matricula, x); acum += adm; return { m, cand: x.length, adm, acum, aptos: n((a) => !a.declinou_em && asoOk(a), x), decl: n((a) => a.declinou_em, x) }; });
  const admAno = (ano) => porMes.filter((x) => x.m.startsWith(ano)).reduce((s, x) => s + x.adm, 0);
  const porLocal = Object.entries(base.reduce((m, a) => { (m[a.local || '—'] = m[a.local || '—'] || []).push(a); return m; }, {})).map(([s, x]) => ({ s, cand: x.length, adm: n((a) => !a.declinou_em && a.matricula, x), proc: n((a) => !a.declinou_em && !a.matricula, x) })).sort((x, y) => y.cand - x.cand);
  const barras = (l, maxv, f1, f2, r1, r2) => `<div class="funil">${l.map((x) => `<div class="l l-dupla"><span>${esc(x.s)}</span><div class="barra dupla" title="${esc(r1)}: ${x[f1]} · ${esc(r2)}: ${x[f2]}"><i style="width:${(x[f1] / maxv) * 100}%"></i><i style="width:${(x[f2] / maxv) * 100}%"></i></div><b class="num">${x[f1]} · ${x[f2]}</b></div>`).join('')}</div>`;
  const maxMes = Math.max(1, ...porMes.map((x) => x.cand));
  const hoje = iso(new Date());
  const prox = (E.turmasCal || []).filter((t) => t.ativo && t.data >= hoje).sort((x, y) => x.data.localeCompare(y.data)).slice(0, 6);
  const viv = base.filter((a) => !a.declinou_em && !a.matricula); const um = viv[0] || {};
  const pendArea = [
    ['Remuneração · validar posição', n((a) => a.remuneracao !== 'ok', viv), 'remuneracao'],
    ['Candidato · enviar documentos', n((a) => ['aguardando', 'pendencia'].includes(a.etapa), viv), 'painel'],
    ['Admissão · conferir documentos', n((a) => a.etapa === 'conferencia' || resumo(a).analise > 0, viv), 'painel'],
    ...('docs_revisados_em' in um ? [['Admissão · revisão final', n((a) => ['docs_ok', 'registro'].includes(a.etapa) && !a.docs_revisados_em, viv), 'painel']] : []),
    ...('soc_cadastrado_em' in um ? [['R&S · cadastrar no SOC', n((a) => !a.soc_cadastrado_em && a.aso_status === 'aguardando', viv), 'saude']] : []),
    ['Saúde · ASO apto', n((a) => !asoOk(a), viv), 'saude'],
    ['R&S · movimentar no LG', n((a) => faltaLg(a), viv), 'painel'],
    ['Admissão · registrar matrícula', n((a) => a.etapa === 'registro', viv), 'painel'],
    ['Admissão · chamado de TI', n((a) => chamadoTiAtrasado(a), base), 'ti'],
    ...('sindicato_status' in um ? [['Relações Sindicais · conferir', n((a) => ['divergente', 'sem_regra'].includes(a.sindicato_status), viv), 'sindicatos']] : []),
  ];
  const maxPend = Math.max(1, ...pendArea.map((x) => x[1]));
  const top3 = porSetor.slice(0, 3); const somaTop = top3.reduce((s, x) => s + x.cand, 0);
  const maisDecl = porSetor.filter((x) => x.decl).sort((x, y) => y.decl - x.decl)[0];
  const maisPend = porSetor.filter((x) => x.pend).sort((x, y) => y.pend - x.pend)[0];
  const pico = porMes.slice().sort((x, y) => y.adm - x.adm)[0];
  const frases = [
    top3.length ? `Concentração: ${top3.map((x) => x.s).join(', ')} somam ${pct(somaTop, cand)}% dos candidatos.` : '',
    maisDecl ? `Desistência: ${maisDecl.s} tem mais declínios (${maisDecl.decl}, ${pct(maisDecl.decl, maisDecl.cand)}% do setor).` : 'Nenhum declínio neste recorte.',
    maisPend ? `Exames: ${maisPend.s} concentra ${pct(maisPend.pend, pendEx)}% dos exames pendentes (${maisPend.pend}).` : 'Nenhum exame pendente neste recorte.',
    pico && pico.adm && !F.mes ? `Evolução: o pico de admissões foi em ${pico.m.slice(5)}/${pico.m.slice(0, 4)} (${pico.adm}).` : '',
    `Remuneração: ${semRem ? plural(semRem, 'candidato ainda sem validação', 'candidatos ainda sem validação') : 'todos validados'}${ocup ? ' · ' + plural(ocup, 'posição ocupada', 'posições ocupadas') : ''}.`,
    cand ? `Perfil: ${pct(n((a) => a.tipo_contrato === 'Estágio'), cand)}% estágio, ${pct(n((a) => /operador|auxiliar/i.test(a.cargo || '')), cand)}% operacional (operador e auxiliar), ${pct(n((a) => a.tipo_contrato === 'Jovem aprendiz'), cand)}% aprendiz.` : '',
  ].filter(Boolean);
  const proxDia = turmas().find((d) => deIso(d) >= deIso(hoje));
  const nEtapa = (id) => n((a) => !a.declinou_em && a.etapa === id);
  const maxE = Math.max(1, ...ETAPAS.map((e) => nEtapa(e.id)));
  const atencao = listaAtencao();
  const alertaSetor = (x) => (x.cand && pct(x.decl, x.cand) >= lim.desist ? chip('c-cr', 'cr', 'Desistência') : x.cand - x.decl && pct(x.pend, x.cand - x.decl) >= lim.exames ? chip('c-at', 'at', 'Exames') : chip('c-ok', 'ok', 'OK'));
  return `<div class="cab"><div class="t"><span class="olho">Visão geral</span><h1>Admissões</h1><p class="apoio">O painel da planilha de admissão (aba RESUMO), calculado pelo sistema a partir do que cada área registra. Filtre por setor, local e mês de início.</p></div>
    <div class="linha-flex">${tem('admissao') ? '<button class="bt" data-act="importar">Importar a planilha atual (.xlsx)</button>' : ''}<button class="bt ama" data-act="ir" data-v="novo">${MI.novo} Cadastrar aprovado</button></div></div>
  ${proxDia ? faixaCorte(proxDia) : ''}
  <div class="filtros"><label class="olho" for="iSetor">Setor</label><select id="iSetor"><option value="">Todos</option>${setores.map((s) => `<option ${s === F.setor ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>
    <label class="olho" for="iLocal">Local</label><select id="iLocal"><option value="">Todos</option>${locais.map((s) => `<option ${s === F.local ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>
    <label class="olho" for="iMes">Mês de início</label><select id="iMes"><option value="">Todos</option>${meses.map((m) => `<option value="${m}" ${m === F.mes ? 'selected' : ''}>${m.slice(5)}/${m.slice(0, 4)}</option>`).join('')}</select>
    ${F.setor || F.local || F.mes ? '<button class="bt peq" data-act="iniLimpar">Limpar filtros</button>' : ''}</div>
  <div class="grade g4">
    ${kpi(admit, 'admitidos (com matrícula)', plural(cand, 'candidato cadastrado', 'candidatos cadastrados'))}
    ${kpi(proc, 'em processo (sem matrícula)', pct(proc, cand) + '% dos candidatos')}
    ${kpi(aptos, 'exames aptos', pct(aptos, cand - decl) + '% de quem segue')}
    ${kpi(pendEx, 'exames pendentes', pct(pendEx, cand - decl) + '% de quem segue', pct(pendEx, cand - decl) >= lim.exames)}
    ${kpi(decl, 'declínios', plural(decl, 'desistência', 'desistências'))}
    ${kpi(pct(decl, cand) + '%', 'de desistência', 'alerta a partir de ' + lim.desist + '%', pct(decl, cand) >= lim.desist)}
    ${kpi(remOk, 'remuneração ok', pct(remOk, cand) + '% validados')}
    ${kpi(ocup, 'vagas ocupadas', 'posição ainda não desocupada', ocup > 0)}
  </div>
  <div class="caixa pilha"><h2>Leitura automática</h2><ul style="margin:0;padding-left:18px">${frases.map((f) => `<li>${esc(f)}</li>`).join('')}</ul></div>
  <div class="grade g2">
    <div class="caixa pilha"><h2>Próximas turmas</h2>${prox.length ? `<div class="tab-wrap"><table><thead><tr><th>Admissão</th><th>Unidades</th><th>Corte</th><th>Pessoas</th><th>Sem matrícula</th></tr></thead><tbody>${prox.map((t) => { const ps = vivas().filter((a) => a.turma_id === t.id); const tarde = Date.now() > new Date(t.corte_inclusao).getTime(); return `<tr><td class="num"><b>${dmd(t.data)}</b></td><td>${esc(t.nome || GRUPO_TURMA[t.grupo])}</td><td>${tarde ? chip('c-ag', '', 'encerrado ' + dmh(t.corte_inclusao)) : chip('c-at', 'at', 'até ' + dmh(t.corte_inclusao))}</td><td class="num">${ps.length}</td><td class="num">${ps.filter((a) => !a.matricula).length}</td></tr>`; }).join('')}</tbody></table></div>` : '<div class="vazio">Nenhuma turma futura no cronograma.</div>'}</div>
    <div class="caixa pilha"><h2>Pendências por área</h2><div class="funil">${pendArea.map(([r, k, v]) => `<div class="l"><span>${navVisivel().some((x) => x[0] === v) ? `<button class="bt-link" data-act="ir" data-v="${v}">${esc(r)}</button>` : esc(r)}</span><div class="barra"><i style="width:${(k / maxPend) * 100}%"></i></div><b class="num">${k}</b></div>`).join('')}</div></div>
  </div>
  <div class="grade g2">
    <div class="caixa pilha"><h2>Onde cada pessoa está</h2><div class="funil">${ETAPAS.map((e) => { const k = nEtapa(e.id); return `<div class="l"><span>${esc(e.n)}</span><div class="barra"><i style="width:${(k / maxE) * 100}%"></i></div><b class="num">${k}</b></div>`; }).join('')}</div></div>
    <div class="caixa pilha"><h2>Precisa de atenção</h2>${atencao.length ? `<div class="tab-wrap"><table><tbody>${atencao.slice(0, 10).map(([p, a, tx]) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.cargo)} · ${esc(a.local)}</span></td><td>${p === 0 ? chip('c-cr', 'cr', tx) : esc(tx)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nada urgente agora.</div>'}</div>
  </div>
  <div class="grade g2">
    <div class="caixa pilha"><h2>Top 10 setores</h2><p class="apoio" style="font-size:12px">admitidos (amarelo) · em processo (azul)</p>${porSetor.length ? barras(porSetor.slice(0, 10).map((x) => ({ s: x.s, adm: x.adm, proc: x.cand - x.adm - x.decl })), Math.max(1, ...porSetor.slice(0, 10).map((x) => Math.max(x.adm, x.cand - x.adm - x.decl))), 'adm', 'proc', 'admitidos', 'em processo') : '<div class="vazio">—</div>'}</div>
    <div class="caixa pilha"><h2>Localidades</h2><p class="apoio" style="font-size:12px">candidatos (amarelo) · admitidos (azul)</p>${porLocal.length ? barras(porLocal.slice(0, 12), Math.max(1, ...porLocal.map((x) => x.cand)), 'cand', 'adm', 'candidatos', 'admitidos') : '<div class="vazio">—</div>'}</div>
  </div>
  <div class="caixa pilha"><h2>Por setor</h2>${porSetor.length ? `<div class="tab-wrap"><table><thead><tr><th>Setor</th><th>Candidatos</th><th>% do total</th><th>Admitidos</th><th>Aptos</th><th>Exames pendentes</th><th>Declínios</th><th>% desistência</th><th>Remun. ok</th><th>Alerta</th></tr></thead><tbody>${porSetor.map((x) => `<tr><td>${esc(x.s)}</td><td class="num">${x.cand}</td><td class="num">${pct(x.cand, cand)}%</td><td class="num">${x.adm}</td><td class="num">${x.aptos}</td><td class="num">${x.pend}</td><td class="num">${x.decl}</td><td class="num">${pct(x.decl, x.cand)}%</td><td class="num">${x.rem}</td><td>${alertaSetor(x)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Sem candidatos neste recorte.</div>'}</div>
  <div class="caixa pilha"><h2>Mês a mês (data de início)</h2>${porMes.length ? `<div class="tab-wrap"><table><thead><tr><th>Mês</th><th>Candidatos</th><th></th><th>Admitidos</th><th>Acumulado</th><th>Aptos</th><th>Declínios</th><th>% desistência</th><th>% do ano</th></tr></thead><tbody>${porMes.map((x) => `<tr><td class="num">${x.m.slice(5)}/${x.m.slice(0, 4)}</td><td class="num">${x.cand}</td><td style="width:26%"><div class="barra-h"><i style="width:${(x.cand / maxMes) * 100}%"></i></div></td><td class="num">${x.adm}</td><td class="num">${x.acum}</td><td class="num">${x.aptos}</td><td class="num">${x.decl}</td><td class="num">${pct(x.decl, x.cand)}%</td><td class="num">${pct(x.adm, admAno(x.m.slice(0, 4)))}%</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Sem datas de início.</div>'}
    <p class="apoio" style="font-size:12px">Admitido = tem matrícula. Exame pendente = ainda sem ASO apto (sem contar declínios). Remuneração ok = posição validada pela Remuneração. Vaga ocupada = a posição ainda não foi desocupada (substituição).</p></div>`;
};
POS.inicio = () => {
  $('#iSetor')?.addEventListener('change', (e) => { E.iniF.setor = e.target.value; render(); });
  $('#iLocal')?.addEventListener('change', (e) => { E.iniF.local = e.target.value; render(); });
  $('#iMes')?.addEventListener('change', (e) => { E.iniF.mes = e.target.value; render(); });
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
// exceção ao corte liberada pelo Time de Admissão para este CPF e esta data (SQL 27)
const soDig = (s) => String(s || '').replace(/\D/g, '');
const excecaoPara = (cpf, data) => (E.excecoes || []).find((x) => x.cpf === soDig(cpf) && x.data_inicio === data && !x.usada_em && !x.cancelada_em);
// Data de admissão (integração): as turmas do cronograma. Depois do corte o R&S só inclui com exceção
// liberada pela Admissão para o CPF; a Admissão inclui sempre (fica registrado).
function opcoesTurma(local, sel, cpf) {
  const ts = turmasParaLocal(local); const livre = tem('admissao'); const trava27 = temSql27();
  const op = ts.map((t) => {
    const tarde = Date.now() > new Date(t.corte_inclusao).getTime(); const exc = tarde && excecaoPara(cpf, t.data);
    const trava = tarde && trava27 && !livre && !exc;
    return `<option value="${t.id}" ${t.id === sel && !trava ? 'selected' : ''} ${trava ? 'disabled' : ''}>${dmd(t.data)} · ${esc(t.nome || GRUPO_TURMA[t.grupo])}${tarde ? (exc ? ' · exceção liberada' : ' · corte encerrado') : ''}</option>`;
  });
  (E.excecoes || []).filter((x) => x.cpf === soDig(cpf) && !x.usada_em && !x.cancelada_em && !ts.some((t) => t.data === x.data_inicio))
    .forEach((x) => op.push(`<option value="data:${x.data_inicio}" ${sel === 'data:' + x.data_inicio ? 'selected' : ''}>${dmd(x.data_inicio)} · exceção liberada pelo Time de Admissão</option>`));
  if (!ts.length && !trava27) { const d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); for (let i = 0; i < 6; i++) { const x = new Date(d.getTime() + i * 7 * DIA); op.push(`<option value="data:${iso(x)}">${dmd(x)} · corte ${dmd(corteDe(iso(x)))} 12h (sem cronograma)</option>`); } }
  if (livre || !trava27) op.push(`<option value="outra" ${sel === 'outra' ? 'selected' : ''}>Outra data, fora do cronograma</option>`);
  if (!op.length) op.push('<option value="">Nenhuma turma aberta: fale com o Time de Admissão</option>');
  return op.join('');
}
const textoCorte = (t) => (t ? `Inclusão até ${dmh(t.corte_inclusao)}${t.corte_remuneracao ? ' · Remuneração valida até ' + dmh(t.corte_remuneracao) : ''} · documentos, ASO e LG até ${dmh(t.corte_documentos)}` : '');
// fretado do formulário: com as linhas do LG grava o código e o nome; sem elas, o texto
function fretadoDoForm(local) {
  const s = $('#n_fretado'); if (!s) return { fretado: null };
  const integ = $('#n_frInteg') && !$('#n_frIntegBox').hidden && $('#n_frInteg').checked;
  const extra = temSql31() ? { fretado_integ: !!integ, fretado_integ_chave: integ ? ($('#n_frIntegSel')?.value || null) : null, fretado_itin_chave: null } : {};
  if (s.tagName === 'SELECT' && s.value.includes('|') && E.itins) {
    const it = E.itins.find((x) => x.chave === s.value);
    if (it) return Object.assign(extra, { fretado: titulo(it.linha), fretado_codigo: it.codigo_lg || null, fretado_itin_chave: it.chave });
  }
  if (Object.keys(extra).length) { const r = fretadoDoFormLg(s); return Object.assign(r, extra, { fretado_itin_chave: null }); }
  return fretadoDoFormLg(s);
}
function fretadoDoFormLg(s) {
  if (s.tagName === 'SELECT' && /^\d+$/.test(s.value) && temSql31()) { const l = (E.tLinhas || []).find((x) => x.codigo_lg === +s.value); return { fretado: l ? nomeLinha(l) : null, fretado_codigo: +s.value }; }
  if (s.tagName === 'SELECT' && s.value.includes('|')) return { fretado: null, fretado_codigo: null };
  const txt = valLista('n_fretado') || null;
  return temSql31() ? { fretado: txt, fretado_codigo: null } : { fretado: txt };
}
const ITENS_TI = ['Notebook', 'Desktop', 'Usuário de rede', 'E-mail', 'Teams', 'SAP', 'Salesforce', 'Celular'];
VIEWS.novo = () => {
  const opt = (arr, sel) => arr.map((v) => `<option ${v === sel ? 'selected' : ''}>${esc(v)}</option>`).join('');
  const deps = ligado('dependentes_ativo'); const l0 = localDe('Pouso Alegre');
  const euAnalista = opcoesLista('analista_rs').find((x) => semAcento(x) === semAcento(E.perfil?.nome)) || '';
  const itensTi = opcoesLista('acesso_ti').length ? opcoesLista('acesso_ti') : ITENS_TI;
  const sinds = [...new Set([...opcoesLista('sindicato'), ...(E.regrasSind || []).map((r) => r.sindicato)])].sort();
  return `<div class="cab"><div class="t"><span class="olho">Atração e Seleção</span><h1>Cadastrar aprovado</h1><p class="apoio">Ao salvar, a Saúde, a Admissão, a Remuneração e Benefícios recebem o aviso, e o candidato recebe a lista de documentos com o prazo do corte da turma. A inclusão fica gravada com o seu nome, a data e a hora. ${temSql27() ? 'Depois do corte da turma o cadastro trava: exceção aprovada, só o Time de Admissão libera.' : ''}</p></div></div>
  <form class="caixa" id="fNovo" novalidate><div class="form">
    <div class="secao-form">Candidato</div>
    <div class="campo l2"><label for="n_nome">Nome completo, como no documento</label><input id="n_nome" autocomplete="off"></div>
    <div class="campo"><label for="n_cpf">CPF</label><input id="n_cpf" inputmode="numeric" placeholder="000.000.000-00"></div>
    <div class="campo"><label for="n_nasc">Data de nascimento</label><input id="n_nasc" type="date"><span class="ajuda">Idade (menor de 18), reservista, dependentes e benefícios.</span></div>
    <div class="campo"><label for="n_tel">WhatsApp</label><input id="n_tel" inputmode="tel" placeholder="(35) 99999-9999"></div>
    <div class="campo l2"><label for="n_email">E-mail pessoal</label><input id="n_email" type="email"><span class="ajuda">Importante para comunicar as próximas etapas (lista de documentos, cobranças e convite). Não é obrigatório: sem ele, o cadastro segue.</span></div>
    <div class="campo"><label for="n_sexo">Sexo</label><select id="n_sexo"><option value="F">Feminino</option><option value="M">Masculino</option></select><span class="ajuda">Define se pede reservista.</span></div>
    <div class="campo"><label for="n_civil">Estado civil</label><select id="n_civil">${opt(['Solteiro(a)', 'Casado(a)', 'União estável', 'Divorciado(a)', 'Viúvo(a)'])}</select><span class="ajuda">Casado(a) ou união estável: certidão, RG e SUS do cônjuge.</span></div>
    <div class="campo"><label for="n_filhos">Filhos</label><input id="n_filhos" type="number" min="0" max="15" value="0" ${deps ? 'readonly title="Vem da lista de dependentes"' : ''}>${deps ? '<span class="ajuda">Contado pela lista de dependentes abaixo.</span>' : ''}</div>
    <div class="campo"><label for="n_tipo">Contratação</label><select id="n_tipo">${opt(['CLT', 'Jovem aprendiz', 'Estágio', 'Temporário'])}</select><span class="ajuda">Estágio vai como S-2300; aprendiz e estágio têm pacote de benefícios próprio.</span></div>
    <div class="campo" id="n_agenciaBox" hidden><label for="n_agencia">Agência</label><input id="n_agencia" placeholder="Agência do temporário"><span class="ajuda">Só para temporário: é contratado pela agência.</span></div>
    ${temSql31() ? '<div class="campo l2"><label>Plano de saúde da Cimed</label><div class="marcas"><label><input type="checkbox" id="n_eradep"> Já é dependente de alguém da Cimed no plano (vai virar titular)</label></div><span class="ajuda">Benefícios move de dependente para titular no WKF da operadora.</span></div>' : ''}
    <div class="campo"><label for="n_pcd">PCD</label><select id="n_pcd">${opt(['Não', 'Auditiva', 'Física', 'Intelectual', 'Visual', 'Mental', 'Reabilitado'])}</select><span class="ajuda">O tipo só a Saúde, a Admissão e a Remuneração veem (cota e laudo).</span></div>
    ${deps ? `<div class="secao-form">Dependentes</div><div class="campo l4">${blocoDeps(E.depsNovo, 'n', true)}</div>` : ''}
    <div class="secao-form">Vaga</div>
    <div class="campo"><label for="n_cargo">Cargo</label><input id="n_cargo"></div>
    <div class="campo"><label for="n_cbo">CBO</label><input id="n_cbo" inputmode="numeric" placeholder="0000-00"><span class="ajuda">Vai na Ordem de Serviço da Segurança.</span></div>
    <div class="campo"><label for="n_setor">Setor</label><input id="n_setor"></div>
    <div class="campo"><label for="n_local">Local de trabalho</label><select id="n_local">${E.locais.map((l) => `<option ${l.nome === 'Pouso Alegre' ? 'selected' : ''}>${esc(l.nome)}</option>`).join('')}</select><span class="ajuda" id="n_localAj"></span></div>
    <div class="campo l2" id="n_motivoCampo"><label>Motivo da vaga</label><div class="marcas" id="n_motivoBox"><label><input type="radio" name="n_motivo" value="aumento"> Aumento de quadro</label><label><input type="radio" name="n_motivo" value="subst"> Substituição</label></div><span class="ajuda">Substituição: a Remuneração confere se a posição já foi desocupada.</span></div>
    <div class="campo l2" id="n_substBox" hidden><label for="n_subst_nome">Pessoa substituída</label><input id="n_subst_nome"></div>
    <div class="campo"><label for="n_posicao">Posição</label><input id="n_posicao" inputmode="numeric"></div>
    <div class="campo"><label for="n_proc">Nº do processo seletivo</label><input id="n_proc" inputmode="numeric"></div>
    <div class="campo l2"><label for="n_analista">Analista responsável (R&amp;S)</label>${selLista('n_analista', 'analista_rs', euAnalista, { vazio: 'Escolha a analista', outra: true, ph: 'Nome da analista' })}<span class="ajuda">Recebe as cobranças da Admissão (LG) e os avisos da Saúde.</span></div>
    <div class="campo l2"><label for="n_req">Requisitante (gestor)</label><input id="n_req"></div>
    <div class="campo l2"><label for="n_req_email">E-mail do requisitante</label><input id="n_req_email" type="email" placeholder="nome@grupocimed.com.br"><span class="ajuda">Para o aviso ao gestor e a cópia do convite.</span></div>
    <div class="campo l2"><label for="n_escala">Escala</label>${selLista('n_escala', 'escala', '', { vazio: 'Escolha a escala', outra: true, ph: '05:30 11:00/12:00 15:18' })}</div>
    <div class="campo"><label for="n_turno">Turno</label>${selLista('n_turno', 'turno', '', { vazio: '—', padrao: ['1º turno', '2º turno', '3º turno', 'Administrativo', 'Comercial'] })}</div>
    <div class="campo"><label for="n_cc_cod">Centro de custo (código)</label><input id="n_cc_cod" inputmode="numeric"><span class="ajuda">Vai na lista da turma (U1/U2) e para a TI.</span></div>
    <div class="campo l2"><label for="n_cc_nome">Centro de custo (nome)</label><input id="n_cc_nome"></div>
    ${temSql27() ? `<div class="campo l2"><label for="n_sind">Sindicato</label><input id="n_sind" list="dlSindN" placeholder="Escolha ou digite"><datalist id="dlSindN">${sinds.map((s) => `<option value="${esc(s)}">`).join('')}</datalist><span class="ajuda" id="n_sindAj"></span></div>` : ''}
    <div class="campo l2"><label>Integração NR-12</label><div class="marcas"><label><input type="checkbox" id="n_nr12"> Opera máquina (NR-12 teórica e prática no 2º dia)</label></div></div>
    <div class="secao-form">Integração</div>
    <div class="campo l2"><label for="n_turma">Data de admissão (integração)</label><select id="n_turma">${opcoesTurma('Pouso Alegre')}</select><span class="ajuda" id="n_turmaAj"></span></div>
    <div class="campo l2"><label for="n_corte">Data corte desta turma</label><input id="n_corte" readonly tabindex="-1"><span class="ajuda">Vem do cronograma do Time de Admissão (tela Cronograma e corte).</span></div>
    <div class="campo" id="n_inicioBox" hidden><label for="n_inicio">Data de início</label><input id="n_inicio" type="date"><span class="ajuda">Fora do cronograma: o DP recebe um alerta.</span></div>
    <div class="campo"><label for="n_dias">Dias de integração</label><input id="n_dias" type="number" min="1" max="10" value="${l0.dias || 1}"><span class="ajuda">Padrão do local; ajuste se precisar.</span></div>
    <div class="campo"><label for="n_hora">Horário da integração</label><input id="n_hora" value="${esc(l0.hora || '08:00')}" placeholder="08:00"></div>
    <div class="campo"><label for="n_formato">Formato</label><select id="n_formato"><option value="">Padrão do local</option><option value="presencial">Presencial</option><option value="online">On-line (Teams)</option></select></div>
    <div class="campo l2"><label for="n_link">Link do Teams ou endereço</label><input id="n_link" placeholder="Padrão da turma, se houver"><span class="ajuda">Vai no convite ao candidato.</span></div>
    <div class="campo"><label for="n_contato">Quem procurar ao chegar</label><input id="n_contato" placeholder="Nome e ramal"></div>
    <div class="secao-form">Para as áreas</div>
    <div class="campo"><label for="n_camisa">Tamanho de camisa</label><select id="n_camisa">${opt(['PP', 'P', 'M', 'G', 'GG', 'XGG', 'Não usa uniforme'], 'M')}</select></div>
    <div class="campo l2"><label for="n_fretado">Linha do fretado${temSql31() ? ' (pelo turno)' : ''}</label><div id="n_fretadoWrap">${selFretado('Pouso Alegre')}</div><span class="ajuda" id="n_fretadoAj">${temSql31() ? 'Só aparecem as linhas que atendem o turno escolhido. O colaborador recebe o itinerário por e-mail.' : 'Benefícios cadastra no Rota para o QR Code.'}</span></div>
    ${temSql31() ? `<div class="campo l2" id="n_frIntegBox"><label>Fretado na integração</label><div class="marcas"><label><input type="checkbox" id="n_frInteg"> Precisará de fretado na integração? (turno administrativo)</label></div><div id="n_frIntegWrap" hidden></div><span class="ajuda">Só Pouso Alegre e Fernão Dias. O e-mail leva o itinerário da integração e o do turno.</span></div>` : ''}
    <div class="campo"><label for="n_cd_entrega">CD de entrega do equipamento</label><input id="n_cd_entrega" placeholder="Só para o comercial"><span class="ajuda">A TI não entrega em casa: vai para o CD mais perto.</span></div>
    <div class="campo l2"><label>Usa veículo da frota?</label><div class="marcas"><label><input type="checkbox" id="n_frota"> Sim, vai dirigir carro da empresa</label><label><input type="checkbox" id="n_motorista"> Motorista profissional (dirigir é a função)</label></div><span class="ajuda">Frotas separa o veículo; motorista profissional pede CNH com EAR e toxicológico.</span></div>
    ${temSql35() ? `<div class="campo l2" id="n_frotaTipoBox" hidden><label for="n_frota_tipo">Uso do carro</label><select id="n_frota_tipo"><option value="">Escolha</option>${Object.entries(FROTA_TIPO).map(([k, x]) => `<option value="${k}">${esc(x)}</option>`).join('')}</select><span class="ajuda" id="n_frotaAj">Frotas recebe o aviso já no cadastro. Prazo da política: 2 dias no comercial, 30 no benefício. A pessoa envia a CNH definitiva e o extrato de pontos.</span></div>` : ''}
    <div class="campo l4"><label>Equipamentos e acessos</label><div class="marcas" id="n_ti">${itensTi.map((x) => `<label><input type="checkbox" value="${esc(x)}"> ${esc(x)}</label>`).join('')}</div><span class="ajuda">A TI recebe a previsão agora; a Admissão abre o chamado com a matrícula usando estes dados.</span></div>
  </div><div class="linha-flex" style="margin-top:18px;justify-content:flex-end"><span class="apoio" id="n_resumo" style="flex:1;min-width:220px"></span><button type="submit" class="bt pri" id="n_salvar">Salvar e avisar as áreas</button></div></form>`;
};
POS.novo = () => {
  const f = $('#fNovo'); let nr12Manual = false, diasManual = false, horaManual = false;
  const turmaAtual = () => E.turmasCal.find((x) => x.id === $('#n_turma').value);
  const dataInicioForm = () => {
    const s = $('#n_turma').value;
    if (s === 'outra') return $('#n_inicio').value || null;
    if (s.startsWith('data:')) return s.slice(5);
    return E.turmasCal.find((x) => x.id === s)?.data || null;
  };
  const upd = () => {
    const l = localDe($('#n_local').value);
    $('#n_localAj').textContent = l.nome ? `Integração em ${l.integ} · calendário ${GRUPO_TURMA[l.grupo_turma] || ''}` : '';
    const t = turmaAtual(); const v = $('#n_turma').value;
    $('#n_inicioBox').hidden = v !== 'outra';
    $('#n_corte').value = t ? textoCorte(t) : v.startsWith('data:') ? 'Data liberada fora do cronograma' : v === 'outra' ? 'Fora do cronograma: o DP recebe um alerta' : '';
    $('#n_turmaAj').textContent = t && Date.now() > new Date(t.corte_inclusao).getTime()
      ? (tem('admissao') ? 'O corte desta turma já passou: a inclusão fica registrada como feita pela Admissão.' : temSql27() ? 'Corte encerrado: inclusão liberada pelo Time de Admissão para este CPF.' : 'O corte de inclusão desta turma já passou: combine com o DP antes de salvar.')
      : temSql27() && !tem('admissao') ? 'Turma com corte encerrado fica bloqueada. Exceção aprovada: peça ao Time de Admissão para liberar o CPF nela.' : '';
    $('#n_substBox').hidden = !f.querySelector('input[name=n_motivo][value=subst]').checked;
    $('#n_agenciaBox').hidden = $('#n_tipo').value !== 'Temporário';
    if ($('#n_frotaTipoBox')) {
      $('#n_frotaTipoBox').hidden = !$('#n_frota').checked;
      const tp = $('#n_frota_tipo').value; const ini = dataInicioForm(); const prazo = frotaPrazo(tp);
      const faltam = ini ? Math.round((new Date(ini + 'T12:00:00') - new Date(new Date().toDateString() + ' 12:00')) / 86400000) : null;
      $('#n_frotaAj').textContent = tp && faltam != null
        ? `Prazo da política: ${prazo} dias. Faltam ${faltam} dias para o início${faltam < prazo ? ': abaixo do prazo, combine a entrega com o gestor e com Frotas.' : '.'}`
        : 'Frotas recebe o aviso já no cadastro. Prazo da política: 2 dias no comercial, 30 no benefício. A pessoa envia a CNH definitiva e o extrato de pontos.';
      $('#n_frotaAj').style.color = tp && faltam != null && faltam < prazo ? 'var(--cr)' : '';
    }
    if (!nr12Manual) $('#n_nr12').checked = l.tipo === 'fabrica' && /operador|produc|produç|máquina|maquina/i.test($('#n_cargo').value);
    if (ligado('dependentes_ativo')) $('#n_filhos').value = E.depsNovo.filter((d) => ['Filho(a)', 'Enteado(a)'].includes(d.parentesco)).length;
    mostrarOutras(f);
    if ($('#n_sindAj')) {
      const sug = sindSugerido($('#n_local').value, $('#n_cargo').value); const inf = $('#n_sind').value.trim();
      $('#n_sindAj').textContent = sug ? 'Regra do local: ' + sug + (inf && semAcento(inf) !== semAcento(sug) ? ' · diferente do informado: Relações Sindicais vai conferir' : '') : 'Sem regra cadastrada para este local: Relações Sindicais confere.';
    }
  };
  const refazTurmas = () => { const sel = $('#n_turma').value; $('#n_turma').innerHTML = opcoesTurma($('#n_local').value, sel, $('#n_cpf').value); };
  $('#n_local').addEventListener('change', () => {
    refazTurmas(); const l = localDe($('#n_local').value);
    if (!diasManual) $('#n_dias').value = l.dias || 1;
    if (!horaManual) $('#n_hora').value = l.hora || '08:00';
    renderFretado();
    // comercial: notebook e celular sempre (reunião com a TI, 15/10/2026)
    if (l.tipo === 'comercial') $('#n_ti').querySelectorAll('input').forEach((c) => { if (['Notebook', 'Celular', 'E-mail', 'Teams', 'Salesforce'].includes(c.value)) c.checked = true; });
    upd();
  });
  $('#n_cpf').addEventListener('change', () => { refazTurmas(); upd(); });
  f.addEventListener('change', (e) => { if (['n_turno', 'n_escala', 'n_escala_outra', 'n_tipo', 'n_frInteg'].includes(e.target.id)) renderFretado(); });
  garantirItins().then(renderFretado);
  $('#n_dias').addEventListener('input', () => { diasManual = true; });
  $('#n_hora').addEventListener('input', () => { horaManual = true; });
  $('#n_nr12').addEventListener('change', () => { nr12Manual = true; });
  f.addEventListener('change', upd); f.addEventListener('input', (e) => { if (['n_cargo', 'n_sind'].includes(e.target.id)) upd(); });
  f.addEventListener('input', (e) => { const el = e.target; if (el.dataset.dep == null) return; const d = E.depsNovo[+el.dataset.dep]; if (!d) return; d[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value; upd(); });
  upd();
  f.addEventListener('submit', async (e) => {
    e.preventDefault();
    f.querySelectorAll('.erro').forEach((x) => x.remove());
    const v = (id) => ($('#' + id)?.value || '').trim(); const erros = [];
    const marca = (id, m) => { const s = document.createElement('span'); s.className = 'erro'; s.textContent = m; $('#' + id).parentElement.appendChild(s); erros.push(id); };
    const motivo = f.querySelector('input[name=n_motivo]:checked')?.value;
    if (v('n_nome').split(/\s+/).length < 2) marca('n_nome', 'Digite o nome completo, como está no documento.');
    if (!cpfValido(v('n_cpf'))) marca('n_cpf', 'CPF inválido. Confira os números.');
    if (v('n_tipo') === 'Temporário' && !v('n_agencia')) marca('n_agencia', 'Informe a agência que contrata o temporário.');
    if (!v('n_nasc')) marca('n_nasc', 'Informe a data de nascimento.');
    if (v('n_email') && !/^\S+@\S+\.\S+$/.test(v('n_email'))) marca('n_email', 'E-mail inválido. Corrija ou deixe em branco.');
    if (v('n_req_email') && !/^\S+@\S+\.\S+$/.test(v('n_req_email'))) marca('n_req_email', 'E-mail inválido.');
    if (v('n_tel').replace(/\D/g, '').length < 10) marca('n_tel', 'Informe o WhatsApp com DDD.');
    ['n_cargo', 'n_req'].forEach((id) => { if (!v(id)) marca(id, 'Obrigatório.'); });
    if (!motivo) marca('n_motivoBox', 'Escolha: aumento de quadro ou substituição.');
    if (opcoesLista('analista_rs').length && !valLista('n_analista')) marca('n_analista', 'Escolha a analista responsável.');
    if ($('#n_frInteg') && !$('#n_frIntegBox').hidden && $('#n_frInteg').checked && !$('#n_frIntegSel')?.value) marca('n_frInteg', 'Escolha a linha do fretado da integração.');
    const turmaSel = v('n_turma'); let turma_id = null, data_inicio = null;
    if (turmaSel === 'outra') { if (!v('n_inicio')) marca('n_inicio', 'Informe a data de início.'); data_inicio = v('n_inicio'); }
    else if (turmaSel.startsWith('data:')) data_inicio = turmaSel.slice(5);
    else { const t = E.turmasCal.find((x) => x.id === turmaSel); turma_id = t?.id || null; data_inicio = t?.data || null; if (!data_inicio) marca('n_turma', 'Escolha a data de admissão.'); }
    const deps = E.depsNovo.filter((d) => (d.nome || '').trim() || d.parentesco);
    deps.forEach((d, i) => {
      if ((d.nome || '').trim().split(/\s+/).length < 2 || !d.parentesco) { erros.push('n_turma'); toast(`Dependente ${i + 1}: nome completo e parentesco são obrigatórios.`, true); }
      if (d.cpf && !cpfValido(d.cpf)) { erros.push('n_turma'); toast(`Dependente ${i + 1}: CPF inválido.`, true); }
    });
    if ($('#n_frotaTipoBox') && $('#n_frota').checked && !v('n_frota_tipo')) marca('n_frota_tipo', 'Escolha o uso do carro: Frotas precisa para o termo e o prazo.');
    if (erros.length) { $('#' + erros[0])?.focus(); return; }
    if (!v('n_email') && !confirm('Sem e-mail, a lista de documentos, as cobranças e o convite da integração não chegam por e-mail ao candidato. O e-mail é importante para comunicar as próximas etapas.\n\nSeguir mesmo assim, sem e-mail?')) { $('#n_email').focus(); return; }
    $('#n_salvar').disabled = true;
    const novo = await tentar(() => api.criar({
      nome: v('n_nome'), cpf: v('n_cpf').replace(/\D/g, ''), nasc: v('n_nasc'), email: v('n_email').toLowerCase() || null, telefone: v('n_tel'),
      sexo: v('n_sexo'), estado_civil: v('n_civil'), filhos: +v('n_filhos') || 0, tipo_contrato: v('n_tipo'), pcd: v('n_pcd'),
      estrangeiro: false, cargo: v('n_cargo'), setor: v('n_setor') || null, local: v('n_local'), data_inicio, turma_id,
      escala: valLista('n_escala') || null, posicao: v('n_posicao') || null, processo: v('n_proc') || null, requisitante: v('n_req'),
      requisitante_email: v('n_req_email').toLowerCase() || null, cc_codigo: v('n_cc_cod') || null, cc_nome: v('n_cc_nome') || null,
      cbo: v('n_cbo') || null, cd_entrega: v('n_cd_entrega') || null, analista_rs: valLista('n_analista') || null,
      dias_integ: +v('n_dias') || null, hora_integ: v('n_hora') || null,
      turno: valLista('n_turno') || null, nr12: $('#n_nr12').checked, substituicao: motivo === 'subst', pessoa_substituida: motivo === 'subst' ? v('n_subst_nome') || null : null,
      formato_integ: v('n_formato') || null, link_ou_endereco: v('n_link') || null, contato_chegada: v('n_contato') || null,
      camisa: v('n_camisa'), ...fretadoDoForm(v('n_local')), usa_frota: $('#n_frota').checked, motorista_profissional: $('#n_motorista').checked,
      agencia: v('n_tipo') === 'Temporário' ? v('n_agencia') || null : null,
      ti_itens: [...document.querySelectorAll('#n_ti input:checked')].map((x) => x.value),
      ...($('#n_sind') ? { sindicato: v('n_sind') || null } : {}),
      ...($('#n_eradep') ? { era_dependente_cimed: $('#n_eradep').checked } : {}),
      ...($('#n_frotaTipoBox') ? { frota_tipo: $('#n_frota').checked ? v('n_frota_tipo') || null : null } : {}),
    }));
    if (novo && deps.length) {
      for (const d of deps) await tentar(() => api.salvarDependente({ admissao_id: novo.id, nome: d.nome.trim(), parentesco: d.parentesco, nasc: d.nasc || null, cpf: (d.cpf || '').replace(/\D/g, '') || null, universitario: !!d.universitario, plano_saude: !!d.plano_saude, odonto: !!d.odonto, salario_familia: !!d.salario_familia }));
    }
    $('#n_salvar') && ($('#n_salvar').disabled = false);
    if (novo) {
      E.depsNovo = [];
      toast('Aprovado cadastrado. ' + (avisadas('cadastro', novo) || 'Nenhuma área com e-mail para avisar ainda.') + (novo.email ? ' A lista de documentos foi para a fila do candidato.' + (novo.fretado_itin_chave || novo.fretado_integ ? ' O itinerário do fretado também.' : '') : ' Sem e-mail: combine o envio dos documentos pelo WhatsApp.'));
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
  if (a.teste) chips.unshift(chip('c-ia', 'info', 'TESTE'));
  return `<button class="cartao ${risco(a) ? 'risco' : ''}" data-act="ficha" data-id="${a.id}"><span class="nm">${esc(titulo(a.nome))}</span><span class="sub">${esc(a.cargo)} · ${esc(a.local)}</span><span class="chips">${chips.join('')}</span><span class="sub num">Integração ${dmd(a.data_inicio)}${a.cobrancas ? ' · ' + plural(a.cobrancas, 'cobrança', 'cobranças') : ''}</span>${a.declinou_em ? '' : fluxoMini(a)}</button>`;
}
function quadro(lista) {
  const decl = lista.filter((a) => a.declinou_em).length;
  return `<div class="quadro-wrap"><div class="quadro">${ETAPAS.map((e) => { const cs = lista.filter((a) => !a.declinou_em && a.etapa === e.id); return `<section class="coluna" aria-label="${esc(e.n)}"><h3>${esc(e.n)}<span class="n num">${cs.length}</span></h3><p class="dica">${esc(e.d)}</p>${cs.map(cartao).join('') || '<div class="vazio-col">Ninguém nesta etapa</div>'}</section>`; }).join('')}</div></div>
  ${decl ? `<p class="apoio" style="font-size:12.5px">${plural(decl, 'pessoa declinou', 'pessoas declinaram')} nesta seleção. Veja na tabela.</p>` : ''}`;
}
function tabela(lista) {
  return `<div class="tab-wrap"><table><thead><tr><th>Nome</th><th>Cargo · setor</th><th>Integração</th><th>R&amp;S</th><th>Etapa</th><th>Fluxo</th><th>Documentos</th><th>ASO</th><th>Remuneração</th><th>Matrícula</th><th>DP</th><th>TI</th><th>Camisa</th><th>Fretado</th></tr></thead><tbody>
  ${lista.map((a) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b>${a.teste ? ' ' + chip('c-ia', 'info', 'TESTE') : ''}</td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.setor || '')}</span></td><td class="num">${dmd(a.data_inicio)}<br><span class="apoio" style="font-size:12px">${esc(a.integ || '')} ${esc(a.hora_integ || '')}</span></td><td>${esc(a.analista_rs || '')}<br><span class="apoio" style="font-size:11.5px">incluído ${quando(a.criado_em)}</span></td><td>${chipEtapa(a)}</td><td>${a.declinou_em ? '—' : fluxoMini(a)}</td><td>${a.declinou_em ? '—' : chipDocs(a)}</td><td>${a.declinou_em ? '—' : chipAso(a)}</td><td>${a.declinou_em ? '—' : chipRemun(a)}</td><td class="num">${esc(a.matricula || '—')}</td><td>${esc(a.analista_dp || '—')}</td><td>${(a.ti_itens || []).length ? esc(a.ti_itens.join(', ')) : '—'}</td><td>${esc(a.camisa || '—')}</td><td>${esc(a.fretado || '—')}</td></tr>`).join('')}
  </tbody></table></div>`;
}

// ---------------------------------------------------------------- fluxo
// Cada etapa com a área responsável e o prazo da turma (agenda da Helaine, 21/10/2026):
// verde feito, vermelho atrasado, amarelo vence em 24h, cinza em aberto.
function etapasFluxo(a) {
  const t = E.turmasCal.find((x) => x.id === a.turma_id);
  const inc = (t && t.corte_inclusao) || a.corte_inclusao || null;
  const rem = (t && t.corte_remuneracao) || a.corte_remuneracao || inc;
  const doc = a.corte_documentos || (a.data_inicio ? corteA(a).toISOString() : null);
  const vespera = (() => { const d = deIso(a.data_inicio); if (!d) return null; do { d.setDate(d.getDate() - 1); } while (d.getDay() % 6 === 0); d.setHours(18, 0, 0, 0); return d.toISOString(); })();
  const dia1 = (() => { const d = deIso(a.data_inicio); if (!d) return null; d.setHours(18, 0, 0, 0); return d.toISOString(); })();
  const ti = prazoChamadoTi(a);
  const e = [
    ['Cadastro', 'R&S', inc, true],
    ['Remuneração', 'Remuneração', rem, a.remuneracao === 'ok'],
  ];
  if ('soc_cadastrado_em' in a) e.push(['Cadastro no SOC', 'R&S', inc, !!a.soc_cadastrado_em || asoOk(a)]);
  if ('saude_contato_em' in a) e.push(['Contato da Saúde', 'Saúde', a.saude_prazo_contato || null, !!a.saude_contato_em || asoOk(a)]);
  e.push(['Documentos', 'Candidato', doc, !!a.avisou_docs_ok]);
  if ('docs_revisados_em' in a) e.push(['Conferência final', 'Admissão' + (a.conferente_docs ? ' · ' + a.conferente_docs.split(' ')[0] : ''), doc, !!a.docs_revisados_em]);
  e.push(['ASO apto', 'Saúde', doc, asoOk(a)]);
  if (ligado('exige_lg_movimentado')) e.push(['LG movimentado', 'R&S', doc, !!a.lg_movimentado_em]);
  e.push(['Matrícula', 'Admissão', vespera, !!a.matricula]);
  if ((a.ti_itens || []).length) e.push(['Chamado de TI', 'Admissão', ti ? new Date(ti.getFullYear(), ti.getMonth(), ti.getDate(), 18).toISOString() : null, !!a.chamado_ti]);
  e.push(['Kit assinado', 'Integração', dia1, a.kit_status === 'concluido']);
  const agora = Date.now();
  return e.map(([n, area, prazo, ok]) => ({ n, area, prazo, st: ok ? 'feito' : prazo && agora > new Date(prazo).getTime() ? 'atrasado' : prazo && new Date(prazo).getTime() - agora < DIA ? 'hoje' : 'aberto' }));
}
const ROT_FLUXO = { feito: 'feito', atrasado: 'atrasado', hoje: 'vence em 24h', aberto: 'em aberto' };
function fluxoVisual(a) {
  return `<div class="fluxo" role="list" aria-label="Fluxo da admissão">${etapasFluxo(a).map((s) => `<div class="fx fx-${s.st}" role="listitem"><b>${esc(s.n)}</b><span>${esc(s.area)}</span><small>${s.st === 'feito' ? 'feito' : s.prazo ? (s.st === 'atrasado' ? 'venceu ' : 'até ') + dmh(s.prazo) : 'sem prazo'}</small></div>`).join('')}</div>`;
}
function fluxoMini(a) {
  const es = etapasFluxo(a);
  return `<span class="fluxo-mini" title="${esc(es.map((s) => s.n + ' (' + s.area + '): ' + ROT_FLUXO[s.st] + (s.prazo && s.st !== 'feito' ? ' · ' + dmh(s.prazo) : '')).join('\n'))}">${es.map((s) => `<i class="fx-${s.st}"></i>`).join('')}</span>`;
}
// ficha: conferência final dos documentos e sindicato (SQL 27)
function fichaConferencia(a) {
  if (!('docs_revisados_em' in a) || a.declinou_em) return '';
  const dp = tem('admissao'); const eu = semAcento(E.perfil?.nome);
  const podeRevisar = tem('admin') || (dp && a.conferente_docs && semAcento(a.conferente_docs) === eu);
  const pronto = ['docs_ok', 'registro', 'assinatura', 'pronto'].includes(a.etapa);
  const sinds = [...new Set([...opcoesLista('sindicato'), ...(E.regrasSind || []).map((r) => r.sindicato)])].sort();
  return `<div class="doc"><div><div class="dn">Conferência final dos documentos</div><div class="dm">${a.docs_revisados_em ? chip('c-ok', 'ok', 'Revisado: tudo ok ' + dm(new Date(a.docs_revisados_em))) + ` <span class="apoio" style="font-size:12px">por ${esc(a.docs_revisados_por || '')}</span>` : chip('c-ag', 'ag', a.conferente_docs ? 'Com ' + a.conferente_docs : 'Sem responsável')}${!a.docs_revisados_em && a.conferente_docs && !pronto ? ' <span class="apoio" style="font-size:12px">O ok final libera quando todos os documentos estiverem conferidos.</span>' : ''}</div></div>
    <div class="acoes">${dp && !a.matricula ? selLista('confSel', 'conferente_docs', a.conferente_docs || '', { vazio: 'Responsável', attrs: ' class="sel-doc"' }) + `<button class="bt peq" data-act="confSalvar" data-id="${a.id}">Salvar</button>` : ''}${podeRevisar && !a.docs_revisados_em ? `<button class="bt peq pri" data-act="docsRevisar" data-id="${a.id}" ${pronto ? '' : 'disabled title="Ainda há documento pendente ou para conferir"'}>Revisei: tudo ok</button>` : ''}${podeRevisar && a.docs_revisados_em && !a.matricula ? `<button class="bt peq fan" data-act="docsRevisarDesfazer" data-id="${a.id}">Desfazer</button>` : ''}</div></div>
  ${'sindicato_status' in a ? `<div class="doc"><div><div class="dn">Sindicato</div><div class="dm">${chipSind(a)} <b>${esc(a.sindicato || 'não informado')}</b>${a.sindicato_sugerido && semAcento(a.sindicato_sugerido) !== semAcento(a.sindicato) ? ` <span class="apoio" style="font-size:12px">· regra do local: ${esc(a.sindicato_sugerido)}</span>` : ''}${a.sindicato_conferido_por ? ` <span class="apoio" style="font-size:12px">· conferido por ${esc(a.sindicato_conferido_por)} em ${quando(a.sindicato_conferido_em)}</span>` : ''}${a.sindicato_obs ? `<div class="motivo">${esc(a.sindicato_obs)}</div>` : ''}</div></div>
    <div class="acoes">${souSindical() && !a.matricula ? `<input id="sindIn" class="sel-doc" list="dlSindF" value="${esc(a.sindicato || a.sindicato_sugerido || '')}" style="width:200px" aria-label="Sindicato"><datalist id="dlSindF">${sinds.map((s) => `<option value="${esc(s)}">`).join('')}</datalist><input id="sindObs" class="sel-doc" style="width:150px" placeholder="observação" aria-label="Observação do sindicato"><button class="bt peq ${['divergente', 'sem_regra'].includes(a.sindicato_status) ? 'pri' : ''}" data-act="sindConferir" data-id="${a.id}">Conferir</button>` : ''}</div></div>` : ''}`;
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
  if ((a.fretado || a.fretado_codigo || a.fretado_itin_chave || a.fretado_integ) && UNID_FRETADO[a.local]) await garantirItins();
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
  const inc = evs.find((h) => /^Cadastrou o aprovado/.test(h.texto || ''));
  const comentario = !!api.comentar && temSql27();
  $('#camada').innerHTML = `<div class="veu" data-act="fechar"><aside class="gaveta" role="dialog" aria-modal="true" aria-label="Ficha" data-pare="1">
   <div class="linha-flex" style="align-items:flex-start"><div class="pilha" style="flex:1;gap:6px"><span class="olho">Ficha do candidato</span><h2 style="font-size:24px">${esc(titulo(a.nome))}</h2><span class="apoio">${esc(a.cargo)}${a.setor ? ' · ' + esc(a.setor) : ''} · ${esc(a.local)}${a.cc_codigo ? ' · CC ' + esc(a.cc_codigo) : ''}</span>
    <div class="linha-flex">${a.teste ? chip('c-ia', 'info', 'TESTE · e-mails só para ' + (E.params?.teste_email || 'a admin')) : ''}${chipEtapa(a)}${foraCorte(a) ? chip('c-cr', 'cr', 'Fora do corte') : ''}${a.pcd !== 'Não' ? chip('c-info', 'info', 'PCD') : ''}${a.tipo_contrato !== 'CLT' ? chip('c-info', 'info', a.tipo_contrato + (a.agencia ? ' · ' + a.agencia : '')) : ''}${a.nasc && idade(a.nasc) < 18 ? chip('c-at', 'at', 'Menor de idade' + (l.tipo === 'fabrica' ? ' em fábrica' : '')) : ''}${a.motorista_profissional ? chip('c-info', 'info', 'Motorista profissional') : ''}${a.formato_integ === 'online' ? chip('c-ag', '', 'Integração on-line') : ''}${a.nr12 ? chip('c-ag', '', 'NR-12') : ''}${a.arquivado_em ? chip('c-ok', 'ok', 'Dossiê arquivado') : ''}${a.origem === 'importacao' ? chip('c-ag', '', 'Veio da planilha') : ''}</div></div>
    <button class="bt" data-act="fechar">Fechar</button></div>
   ${inc ? `<p class="apoio" style="font-size:12.5px;margin:0">Incluído por <b>${esc(a.incluido_por || inc.quem)}</b> em ${quando(inc.em)}${a.origem === 'importacao' ? ' (importação da planilha)' : ''}${a.analista_rs ? ' · analista responsável: ' + esc(a.analista_rs) : ''}</p>` : ''}
   ${!a.declinou_em ? fluxoVisual(a) : ''}
   ${a.declinou_em ? `<div class="faixa"><p><b>Declinou</b> em ${quando(a.declinou_em)}: ${esc(a.declinou_motivo || '')}.${a.declinou_sensivel ? ' <span class="apoio">Motivo sensível: as áreas receberam só o aviso de cancelamento.</span>' : ''}</p></div>` : ''}
   <div class="grade g2">
    <div class="caixa pilha" style="gap:8px"><h3>Integração</h3><span class="num">${dmd(a.data_inicio)} às ${esc(a.hora_integ || '')}</span><span class="apoio">${esc(a.integ || '')} · ${plural(a.dias_integ || 1, 'dia', 'dias')} · ${a.formato_integ === 'online' ? 'on-line (Teams)' : 'presencial'}${a.nr12 ? ' · NR-12' : ''}</span>
     ${a.link_ou_endereco ? `<span class="apoio">${a.formato_integ === 'online' ? 'Link' : 'Endereço'}: ${esc(a.link_ou_endereco)}</span>` : ''}${a.contato_chegada ? `<span class="apoio">Ao chegar, procurar: ${esc(a.contato_chegada)}</span>` : ''}
     <span class="apoio">Turma: ${turmaCal ? esc(turmaCal.nome || GRUPO_TURMA[turmaCal.grupo]) + ' · inclusão até ' + dmh(turmaCal.corte_inclusao) : 'fora do cronograma'}</span><span class="apoio">Corte de documentos: ${dmh(corteA(a))}</span>
     ${podeRs && !a.declinou_em && !a.matricula ? `<button class="bt peq" data-act="turma" data-id="${a.id}" data-t="${prox}" data-tid="${proxTurma ? proxTurma.id : ''}">Mover para a turma de ${dm(prox)}</button>` : ''}</div>
    <div class="caixa pilha" style="gap:8px"><h3>Contato</h3><span>${esc(a.telefone || '—')}</span><span>${esc(a.email || 'sem e-mail')}</span>${tem('saude', 'admissao', 'remuneracao') && a.cpf ? `<span class="apoio num">CPF ${fmtCpf(a.cpf)}${a.nasc ? ' · ' + idade(a.nasc) + ' anos' : ''}</span>` : ''}<span class="apoio">Gestor: ${esc(a.requisitante || '—')}${a.requisitante_email ? ' · ' + esc(a.requisitante_email) : ' · <b>sem e-mail</b> (o aviso ao gestor não sai)'}</span></div>
   </div>
   <div class="caixa pilha"><h3>Checklist da admissão</h3><div class="docs">
    <div class="doc"><div><div class="dn">Exame admissional (ASO)</div><div class="dm">${chipAso(a)}${a.aso_obs ? ' ' + chip('c-info', 'info', a.aso_obs) : ''}${a.aso_restricao ? ` <span class="apoio" style="font-size:12px">Restrição: ${esc(a.aso_restricao)}</span>` : ''}${a.aso_responsavel ? ` <span class="apoio" style="font-size:12px">agendamento: ${esc(a.aso_responsavel)}</span>` : ''}${a.aso_status === 'inapto' && !a.declinou_em ? '<div class="motivo">Inapto: a Saúde decide aqui se o exame é refeito (preencha data, hora e local e clique em Remarcar) ou se a admissão é cancelada. A matrícula fica bloqueada. R&S e Admissão já foram avisados.</div>' : ''}</div></div><div class="acoes">${podeAso && !a.declinou_em ? `<select class="sel-doc" id="asoSel"><option value="aguardando" ${a.aso_status === 'aguardando' ? 'selected' : ''}>A agendar</option><option value="agendado" ${a.aso_status === 'agendado' ? 'selected' : ''}>Agendado</option><option value="apto" ${a.aso_status === 'apto' ? 'selected' : ''}>Apto</option><option value="apto_restricao" ${a.aso_status === 'apto_restricao' ? 'selected' : ''}>Apto com restrição</option><option value="inapto" ${a.aso_status === 'inapto' ? 'selected' : ''}>Inapto</option></select><input type="date" id="asoData" class="sel-doc" value="${a.aso_data || ''}" aria-label="Data do ASO"><input id="asoHora" class="sel-doc" style="width:70px" value="${esc(a.aso_hora || '')}" placeholder="hora" aria-label="Hora do exame"><input id="asoLocal" class="sel-doc" style="width:150px" value="${esc(a.aso_local || '')}" placeholder="ambulatório ou clínica" aria-label="Local do exame"><input id="asoRestr" class="sel-doc" style="width:150px" value="${esc(a.aso_restricao || '')}" placeholder="restrição, se houver" aria-label="Restrição">${'aso_obs' in a ? selLista('asoObs', 'aso_obs', a.aso_obs || '', { vazio: 'Observação', attrs: ' class="sel-doc"' }) + selLista('asoResp', 'agendador_aso', a.aso_responsavel || '', { vazio: 'Quem agendou', attrs: ' class="sel-doc"', ph: 'quem agendou' }) : ''}<button class="bt peq" data-act="salvarAso" data-id="${a.id}">Salvar</button>${a.aso_status === 'inapto' ? `<button class="bt peq pri" data-act="asoRemarcar" data-id="${a.id}" title="Exame refeito: volta para agendado com a data informada">Remarcar exame</button><button class="bt peq fan" data-act="asoCancelar" data-id="${a.id}" title="Cancela a admissão por ASO inapto; as áreas recebem só o aviso de cancelamento">Cancelar admissão</button>` : ''}` : ''}</div></div>
    ${fichaSaude(a)}
    ${fichaConferencia(a)}
    ${temOs(a) ? `<div class="doc"><div><div class="dn">Ordem de Serviço · Segurança do Trabalho</div><div class="dm">${chipOs(a)}${a.cbo ? ' ' + chip('c-ag', '', 'CBO ' + a.cbo) : ''}${a.os_validada_por ? ` <span class="apoio" style="font-size:12px">validada por ${esc(a.os_validada_por)}</span>` : ''}${a.os_status !== 'validada' ? '<div class="motivo">A Esteira só gera a OS no kit com os campos preenchidos. A Segurança valida na tela Segurança do Trabalho.</div>' : ''}</div></div><div class="acoes">${souSeguranca() && !a.declinou_em ? `<button class="bt peq" data-act="osAbrir" data-id="${a.id}">Editar OS</button>` : ''}</div></div>` : ''}
    ${exigeLg || a.lg_movimentado_em ? `<div class="doc"><div><div class="dn">Movimentado no LG (recrutadora)</div><div class="dm">${chipLg}${a.lg_movimentado_por ? ` <span class="apoio" style="font-size:12px">por ${esc(a.lg_movimentado_por)}</span>` : ''}${!a.lg_movimentado_em && a.lg_avisado_em ? ` <span class="apoio" style="font-size:12px">aviso à analista em ${quando(a.lg_avisado_em)}</span>` : ''}${!a.lg_movimentado_em && a.lg_cobrado_em ? ` <span class="apoio" style="font-size:12px">· cobrada em ${quando(a.lg_cobrado_em)}${a.lg_cobrancas > 1 ? ' (' + a.lg_cobrancas + ' vezes)' : ''}</span>` : ''}${!a.lg_movimentado_em && exigeLg ? '<div class="motivo">Sem a movimentação para a etapa "Admissão" no LG, o DP não registra.</div>' : ''}</div></div><div class="acoes">${podeRs && !a.declinou_em && !a.matricula ? (a.lg_movimentado_em ? `<button class="bt peq fan" data-act="lgDesfazer" data-id="${a.id}">Desfazer</button>` : `<button class="bt peq pri" data-act="lgMovimentado" data-id="${a.id}">Movimentei no LG</button>`) : ''}${dp && 'lg_cobrancas' in a && !a.declinou_em && !a.matricula && !a.lg_movimentado_em && asoOk(a) ? `<button class="bt peq" data-act="cobrarLg" data-id="${a.id}" title="E-mail à analista de R&S responsável, e o sino dela acende">Cobrar a recrutadora</button>` : ''}</div></div>` : ''}
    <div class="doc"><div><div class="dn">Remuneração · posição ${esc(a.posicao || '—')}</div><div class="dm">${chipRemun(a)}</div></div><div class="acoes">${podeRem && !a.declinou_em && a.remuneracao !== 'ok' ? `<button class="bt peq" data-act="remun" data-id="${a.id}" data-v="ok">Validada</button><button class="bt peq" data-act="remun" data-id="${a.id}" data-v="ocupada">Posição ocupada</button>` : ''}</div></div>
    <div class="doc"><div><div class="dn">${a.tipo_contrato === 'Estágio' ? 'Cadastro do estágio e eSocial (S-2300)' : 'Matrícula e eSocial'}</div><div class="dm">${a.matricula ? chip('c-esc', '', a.matricula) + (a.esocial_enviado_em ? ' ' + chip('c-ok', 'ok', 'eSocial ' + dm(a.esocial_enviado_em) + (a.esocial_evento ? ' · ' + a.esocial_evento.split(' ')[0] : '')) : ' ' + chip('c-ag', 'ag', 'eSocial não registrado')) + (a.chamado_ti ? ' ' + chip('c-ag', '', 'Chamado TI ' + a.chamado_ti) : '') + ` <span class="apoio" style="font-size:12px">${esc(avisadas('matricula', a) || 'Nenhuma área com e-mail cadastrado para este momento.')}</span>` : a.etapa === 'registro' ? 'Pronto para registrar no LG' : 'Libera com documentos, ASO apto, remuneração ok' + (exigeLg ? ' e LG movimentado' : '')}</div></div><div class="acoes">${dp && a.etapa === 'registro' ? `<input id="matIn" class="sel-doc" style="width:110px" inputmode="numeric" placeholder="Matrícula" aria-label="Matrícula"><button class="bt peq pri" data-act="matricular" data-id="${a.id}">Registrar</button>` : ''}${dp && a.matricula && !a.declinou_em ? `<input type="date" id="esocialData" class="sel-doc" value="${a.esocial_enviado_em || ''}" aria-label="eSocial enviado em"><select id="esocialEv" class="sel-doc" aria-label="Evento"><option value="">Evento</option>${ESOCIAL.map((x) => `<option ${x === a.esocial_evento ? 'selected' : ''}>${x}</option>`).join('')}</select>${inputFicha('chamadoIn', 'Nº do chamado TI', `value="${esc(a.chamado_ti || '')}" style="width:130px"`)}<button class="bt peq" data-act="salvarRegistro" data-id="${a.id}">Salvar</button>${(a.ti_itens || []).length ? `<button class="bt peq" data-act="copiarChamado" data-id="${a.id}" title="Texto pronto para abrir o chamado no Jira">Copiar texto do chamado</button>` : ''}` : ''}</div></div>
    <div class="doc"><div><div class="dn">Para as áreas</div><div class="dm">Camisa ${esc(a.camisa || '—')} · ${a.fretado ? 'Fretado ' + esc(a.fretado) : 'Sem fretado'} · ${(a.ti_itens || []).length ? 'TI: ' + esc(a.ti_itens.join(', ')) + (a.cd_entrega ? ' (entrega no ' + esc(a.cd_entrega) + ')' : '') + (prazoChamadoTi(a) && !a.chamado_ti && !a.declinou_em ? ' · ' + chip(chamadoTiAtrasado(a) ? 'c-cr' : 'c-at', chamadoTiAtrasado(a) ? 'cr' : 'at', 'chamado TI até ' + dm(prazoChamadoTi(a))) : '') : 'Sem equipamento'}${a.usa_frota ? ' · Frota' + (a.frota_tipo ? ' (' + esc(FROTA_TIPO[a.frota_tipo] || a.frota_tipo) + ')' : '') : ''}${a.turno ? ' · ' + esc(a.turno) : ''}${a.cc_codigo || a.cc_nome ? ' · CC ' + esc([a.cc_codigo, a.cc_nome].filter(Boolean).join(' ')) : ''}${a.substituicao ? ' · Substitui ' + esc(a.pessoa_substituida || '(não informado)') : ''}</div></div></div>
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
   ${(a.fretado || a.fretado_codigo || a.fretado_itin_chave || a.fretado_integ) && UNID_FRETADO[a.local] && temSql31() ? `<div class="caixa pilha" style="gap:6px"><h3>Fretado · itinerários</h3>${blocoItinerario(a)}</div>` : ''}
   ${dp && (E.acessosFicha || []).length ? `<div class="caixa pilha" style="gap:6px"><h3>Quem abriu documentos</h3>${E.acessosFicha.map((x) => `<span class="apoio" style="font-size:12.5px">${esc((docs.find((d) => d.tipo === x.tipo) || {}).nome || x.tipo)} · ${esc(x.quem)} · ${quando(x.em)}</span>`).join('')}</div>` : ''}
   <div class="caixa pilha"><h3>Linha do tempo</h3><div class="tempo">${evs.map((h) => `<div class="ev"><span class="pt"></span><div>${esc(textoEvento(h.texto))}<small>${quando(h.em)} · ${esc(h.quem)}</small></div></div>`).join('') || '<span class="apoio">Sem registros.</span>'}</div>
    ${tem('rs', 'admissao', 'saude', 'remuneracao') ? `<form id="fObs" class="linha-flex"><input id="obsTx" maxlength="1000" placeholder="${comentario ? 'Comentar · escreva @Nome Sobrenome para avisar alguém (acende o sino dela)' : 'Escrever uma observação (fica com seu nome e a hora)'}" style="flex:1;min-width:200px;border:1px solid var(--linha);border-radius:10px;padding:9px 12px;background:var(--sup)"><button class="bt" type="submit">${comentario ? 'Comentar' : 'Anotar'}</button></form>` : ''}</div>
   <div class="caixa pilha"><h3>Dados da vaga</h3><div class="grade g2" style="gap:6px 16px;font-size:13px">${[['Escala', a.escala], ['Turno', a.turno], ['Posição', a.posicao], ['Processo seletivo', a.processo], ['Centro de custo', [a.cc_codigo, a.cc_nome].filter(Boolean).join(' · ')], ['Vaga', a.substituicao ? 'Substituição de ' + (a.pessoa_substituida || '?') : 'Aumento de quadro'], ['Requisitante', a.requisitante], ['E-mail do requisitante', a.requisitante_email], ['Analista R&S', a.analista_rs], ['Analista DP', a.analista_dp], ...('sindicato' in a ? [['Sindicato', a.sindicato]] : []), ['Dias de integração', a.dias_integ], ['Horário da integração', a.hora_integ], ['CBO', a.cbo], ['Estado civil', a.estado_civil], ['Filhos', a.filhos], ['Formato da integração', a.formato_integ === 'online' ? 'On-line' : 'Presencial'], ['Chamado de TI', a.chamado_ti]].map(([k, v]) => `<div><span class="apoio">${esc(k)}:</span> <b>${esc(v || '—')}</b></div>`).join('')}</div>
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
  $('#fObs')?.addEventListener('submit', async (e) => { e.preventDefault(); const t = $('#obsTx').value.trim(); if (!t) return; await tentar(() => (api.comentar && temSql27() ? api.comentar(a.id, t) : api.anotar(a.id, t)), /@\S/.test(t) && temSql27() ? 'Comentário registrado; quem foi mencionado recebe no sino' : 'Registrado'); abrirFicha(a.id); });
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
function perguntar(tit, dica, padrao = 'A analista pediu um novo envio') {
  return new Promise((ok) => {
    const div = document.createElement('div'); div.className = 'modal-veu';
    div.innerHTML = `<form class="modal" role="dialog" aria-modal="true"><h2>${esc(tit)}</h2><input id="pTx" class="sel-doc" style="min-height:42px" placeholder="${esc(dica)}"><div class="linha-flex" style="justify-content:flex-end"><button type="button" class="bt" id="pNao">Cancelar</button><button class="bt pri">Confirmar</button></div></form>`;
    document.body.appendChild(div); div.querySelector('#pTx').focus();
    div.querySelector('#pNao').onclick = () => { div.remove(); ok(null); };
    div.querySelector('form').onsubmit = (e) => { e.preventDefault(); const v = div.querySelector('#pTx').value.trim(); div.remove(); ok(v || padrao); };
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
// ficha: fluxo da Saúde (SOC → contato e vacinas → exame → parecer) e o que foi ao R&S
function fichaSaude(a) {
  if (!fluxoSaude(a) || a.declinou_em || a.matricula || !tem('rs', 'admissao', 'saude') || (asoOk(a) && !a.soc_cadastrado_em)) return '';
  const saude = tem('saude', 'admissao'); const rs = tem('rs', 'admissao');
  const vacs = (E.params?.vacinas_obrigatorias || '').split(';').map((x) => x.trim()).filter(Boolean);
  const c = examesCompl(a);
  const linhas = [];
  linhas.push(`<div class="doc"><div><div class="dn">Cadastro no SOC</div><div class="dm">${chipSoc(a)}${a.soc_cadastrado_em ? ` <span class="apoio" style="font-size:12px">às ${hm(new Date(a.soc_cadastrado_em))}${a.soc_cadastrado_por ? ' · ' + esc(a.soc_cadastrado_por) : ''} · contato da Saúde até ${dmh(a.saude_prazo_contato)}</span>` : ' <span class="apoio" style="font-size:12px">Ao marcar, o candidato recebe o e-mail da etapa de exame médico e começa a contar o prazo de contato da Saúde.</span>'}${a.soc_erro ? `<div class="motivo">Erro apontado pela Saúde ${a.soc_erro_em ? 'em ' + quando(a.soc_erro_em) : ''}: ${esc(a.soc_erro)}. ${rs ? 'Corrija no SOC e marque "SOC corrigido".' : 'Aguardando o R&S corrigir.'}</div>` : ''}</div></div>
    <div class="acoes">${!a.soc_cadastrado_em ? `<button class="bt peq pri" data-act="socMarcar" data-id="${a.id}">Marcar cadastrado no SOC</button>` : asoOk(a) ? '' : `<button class="bt peq" data-act="socDesmarcar" data-id="${a.id}" title="Desfaz a marcação (o e-mail ao candidato já enviado não volta)">Desmarcar</button>`}${a.soc_erro && rs ? `<button class="bt peq pri" data-act="rsSaude" data-d="soc_corrigido" data-id="${a.id}">SOC corrigido</button>` : ''}</div></div>`);
  if (asoOk(a)) return linhas.join('');
  if (a.soc_cadastrado_em) {
    linhas.push(`<div class="doc"><div><div class="dn">Contato e vacinas (Saúde)</div><div class="dm">${chipContato(a)} ${chipVacina(a)}<div class="apoio" style="font-size:12px;margin-top:4px">Vacinas obrigatórias: ${esc(vacs.join(', ') || 'não definidas')}${E.params?.vacina_bloqueia_apto !== 'nao' ? ' · sem a comprovação o ASO apto não grava' : ''}${a.vacina_obs ? ' · ' + esc(a.vacina_obs) : ''}</div></div></div>
    <div class="acoes">${saude && !a.saude_escalado_em ? `${a.saude_contato_em && !a.saude_sem_contato_desde ? '' : `<button class="bt peq pri" data-act="saudeContato" data-id="${a.id}">Contato feito</button>`}${a.saude_sem_contato_desde ? '' : `<button class="bt peq" data-act="saudeSemContato" data-id="${a.id}" title="Começa a janela de ${esc(E.params?.saude_janela_contato_horas || '24')}h de tentativas">Não atende</button>`}
      <select class="sel-doc" id="vacSel"><option value="pendente" ${a.vacina_status === 'pendente' ? 'selected' : ''}>Vacinas pendentes</option><option value="comprovada" ${a.vacina_status === 'comprovada' ? 'selected' : ''}>Vacinas comprovadas</option><option value="recusada" ${a.vacina_status === 'recusada' ? 'selected' : ''}>Recusou a vacinação</option></select><input id="vacObs" class="sel-doc" style="width:160px" value="${esc(a.vacina_obs || '')}" placeholder="observação" aria-label="Observação da vacina"><button class="bt peq" data-act="saudeVacina" data-id="${a.id}">Salvar vacinas</button>` : ''}</div></div>`);
    linhas.push(`<div class="doc"><div><div class="dn">Exames da função e clínica</div><div class="dm"><span>${a.exames_tabela ? esc(a.exames_tabela) : '<span class="apoio">Função não encontrada na tabela de exames: a Saúde define abaixo.</span>'}</span><div class="apoio" style="font-size:12px;margin-top:4px">Prazo do ASO: ${c ? '10' : '5'} dias úteis${a.saude_prazo_aso ? ', até ' + dm(a.saude_prazo_aso) : ''}${prazoAsoPassa(a) ? ' (passa da integração de ' + dm(a.data_inicio) + ')' : ''}${a.exames_complementares != null ? ' · ajustado pela Saúde' : ''} · ${plural(a.aso_remarcacoes || 0, 'remarcação', 'remarcações')} de ${remarcMax(a)} permitida${remarcMax(a) > 1 ? 's' : ''}${a.aso_modo ? ' · ' + esc(MODO_ASO[a.aso_modo]) + (a.aso_clinica ? ': ' + esc(a.aso_clinica) : '') : ''}</div></div></div>
    <div class="acoes">${saude && !a.saude_escalado_em ? `<select class="sel-doc" id="complSel" aria-label="Exames complementares"><option value="" ${a.exames_complementares == null ? 'selected' : ''}>Conforme a tabela</option><option value="sim" ${a.exames_complementares === true ? 'selected' : ''}>Com complementares (10 dias)</option><option value="nao" ${a.exames_complementares === false ? 'selected' : ''}>Só clínico (5 dias)</option></select><button class="bt peq" data-act="saudeCompl" data-id="${a.id}">Salvar</button>
      <select class="sel-doc" id="modoSel" aria-label="Onde faz o exame"><option value="">Onde faz o exame</option>${Object.entries(MODO_ASO).map(([k, v]) => `<option value="${k}" ${a.aso_modo === k ? 'selected' : ''}>${v}</option>`).join('')}</select><input id="modoClinica" class="sel-doc" style="width:160px" value="${esc(a.aso_clinica || '')}" placeholder="clínica e cidade" aria-label="Clínica"><button class="bt peq" data-act="saudeModo" data-id="${a.id}">Salvar</button>${a.aso_status === 'agendado' && !a.aso_faltou_em ? `<button class="bt peq fan" data-act="saudeFalta" data-id="${a.id}">Faltou ao exame</button>` : ''}` : ''}</div></div>`);
  }
  if (a.saude_escalado_em) {
    linhas.push(`<div class="doc"><div><div class="dn">Enviado ao R&S</div><div class="dm"><div class="motivo">A Saúde enviou em ${quando(a.saude_escalado_em)}: ${esc(MOTIVO_SAUDE[a.saude_escalado_motivo] || a.saude_escalado_motivo || '')}.${a.saude_escalado_motivo === 'remarcacao_excedida' ? ' Custos de novo agendamento por falha do candidato recaem sobre a área solicitante.' : ''} ${a.saude_escalado_motivo === 'erro_soc' ? (rs ? 'Corrija o cadastro no SOC e clique em "SOC corrigido" acima.' : 'Aguardando o R&S corrigir o SOC.') : rs ? 'Decida: nova tentativa (a Saúde retoma o exame) ou declinar a vaga.' : 'Aguardando a decisão do R&S.'}</div></div></div>
    <div class="acoes">${rs && a.saude_escalado_motivo !== 'erro_soc' ? `<button class="bt peq pri" data-act="rsSaude" data-d="nova_tentativa" data-id="${a.id}">Nova tentativa</button><button class="bt peq fan" data-act="rsSaude" data-d="declinar" data-id="${a.id}">Declinar a vaga</button>` : ''}</div></div>`);
  } else if (saude && a.soc_cadastrado_em) {
    const sugestao = (a.aso_remarcacoes || 0) >= remarcMax(a) ? 'remarcacao_excedida' : a.aso_faltou_em ? 'falta' : 'sem_contato';
    linhas.push(`<div class="doc"><div><div class="dn">Enviar ao R&S</div><div class="dm apoio" style="font-size:12px">Sem contato ou falta depois de ${esc(E.params?.saude_janela_contato_horas || '24')}h de tentativas, ou nova remarcação além do limite: o R&S decide se segue ou declina. Erro no cadastro do SOC: o R&S corrige.${janelaContatoVencida(a) ? ' <b>A janela de tentativas já passou.</b>' : ''}</div></div>
    <div class="acoes"><select class="sel-doc" id="escMotivo" aria-label="Motivo"><option value="sem_contato" ${sugestao === 'sem_contato' ? 'selected' : ''}>Sem contato há 24h</option><option value="falta" ${sugestao === 'falta' ? 'selected' : ''}>Faltou ao exame</option><option value="remarcacao_excedida" ${sugestao === 'remarcacao_excedida' ? 'selected' : ''}>Precisa remarcar de novo</option></select><input id="escTexto" class="sel-doc" style="width:150px" placeholder="detalhe (opcional)" aria-label="Detalhe"><button class="bt peq fan" data-act="saudeEscalar" data-id="${a.id}">Enviar ao R&S</button>
      <input id="socErro" class="sel-doc" style="width:170px" placeholder="o que está errado no SOC" aria-label="Erro no SOC"><button class="bt peq" data-act="saudeErroSoc" data-id="${a.id}">Apontar erro no SOC</button></div></div>`);
  }
  return linhas.join('');
}
function filaSaude() {
  const l = vivas().filter((a) => !a.matricula && a.kit_status !== 'concluido').sort((x, y) => (x.data_inicio + (asoOk(x) ? '1' : '0') + x.local).localeCompare(y.data_inicio + (asoOk(y) ? '1' : '0') + y.local));
  if (!l.length) return '<div class="vazio">Ninguém esperando exame.</div>';
  const corte = (a) => a.corte_documentos ? Math.ceil((new Date(a.corte_documentos).getTime() - Date.now()) / 86400000) : null;
  const fx = l.some(fluxoSaude);
  return `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Cargo · local</th><th>Integração</th>${fx ? '<th>SOC</th><th>Contato</th><th>Vacinas</th>' : '<th>Contato</th>'}<th>Exame</th>${fx ? '<th>Prazo do ASO</th>' : ''}<th>Dias até o corte</th><th>LG</th></tr></thead><tbody>
  ${l.map((a) => `<tr class="clic" data-act="ficha" data-id="${a.id}"><td><b>${esc(titulo(a.nome))}</b>${a.pcd !== 'Não' ? ' ' + chip('c-info', 'info', 'PCD') : ''}<br><span class="apoio num" style="font-size:12px">${esc(a.telefone || '')}${a.cpf ? (a.telefone ? ' · ' : '') + 'CPF ' + fmtCpf(a.cpf) : ''}</span></td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.local)}${a.setor ? ' · ' + esc(a.setor) : ''}</span></td><td class="num">${dmd(a.data_inicio)}</td>${fx ? `<td>${chipSoc(a)}</td><td>${asoOk(a) ? '—' : chipContato(a)}</td><td>${chipVacina(a)}</td>` : `<td>${esc(a.email || '—')}</td>`}<td>${chipAso(a)}${a.aso_modo ? '<br><span class="apoio" style="font-size:12px">' + esc(MODO_ASO[a.aso_modo]) + (a.aso_clinica ? ' · ' + esc(a.aso_clinica) : '') + '</span>' : ''}${a.aso_remarcacoes ? '<br><span class="apoio" style="font-size:12px">' + plural(a.aso_remarcacoes, 'remarcação', 'remarcações') + ' de ' + remarcMax(a) + '</span>' : ''}${a.aso_restricao ? '<br><span class="apoio" style="font-size:12px">' + esc(a.aso_restricao) + '</span>' : ''}</td>${fx ? `<td>${asoOk(a) ? '—' : chipPrazoAso(a)}</td>` : ''}<td class="num">${corte(a) == null ? '—' : corte(a) < 0 ? chip('c-cr', 'cr', 'corte passou') : corte(a) <= 2 && !asoOk(a) ? chip('c-at', 'at', corte(a) + ' d') : corte(a) + ' d'}</td><td>${asoOk(a) ? (a.lg_movimentado_em ? chip('c-ok', 'ok', 'LG ' + dm(new Date(a.lg_movimentado_em))) : chip('c-ag', 'ag', 'LG pendente')) : '—'}</td></tr>`).join('')}
  </tbody></table></div>`;
}
// parâmetros que a Saúde ajusta (SQL 26): a lista de vacinas e se ela bloqueia o apto
function vacinasSaude() {
  const pode = tem('saude', 'admissao', 'admin');
  if (E.params?.vacinas_obrigatorias == null) return '';
  return `<div class="caixa pilha"><h2>Vacinas obrigatórias</h2><p class="apoio">A lista vai no e-mail da etapa de exame médico e aparece na ficha. Separe as vacinas por ponto e vírgula. Com o bloqueio ligado, o ASO apto só grava depois de a vacinação ser registrada como comprovada.</p>
   <div class="linha-flex" style="align-items:flex-end;flex-wrap:wrap"><div class="campo" style="flex:1;min-width:240px"><label for="vacLista">Vacinas</label><input id="vacLista" value="${esc(E.params?.vacinas_obrigatorias || '')}" ${pode ? '' : 'readonly'}></div>
   <div class="campo"><label for="vacBloq">Bloqueia o apto</label><select id="vacBloq" class="sel-doc" ${pode ? '' : 'disabled'}><option value="sim" ${E.params?.vacina_bloqueia_apto !== 'nao' ? 'selected' : ''}>Sim</option><option value="nao" ${E.params?.vacina_bloqueia_apto === 'nao' ? 'selected' : ''}>Não, só registra</option></select></div>
   ${pode ? '<button class="bt pri" data-act="saudeParam">Salvar</button>' : ''}</div></div>`;
}
VIEWS.saude = () => `<div class="cab"><div class="t"><span class="olho">Saúde ocupacional</span><h1>Exame admissional</h1><p class="apoio">${E.params?.vacinas_obrigatorias == null ? 'Quem precisa de exame, por turma. Clique na pessoa para agendar (data, hora e local: o candidato recebe o e-mail), registrar apto, apto com restrição ou inapto. Apto libera a matrícula e pede à recrutadora a movimentação no LG; inapto avisa o R&S e a Admissão, e a Saúde decide na ficha se remarca o exame ou cancela a admissão. O ASO oficial continua no SOC.' : `Fluxo da Saúde: o R&S cadastra no SOC; cadastro até 12h, a Saúde faz o contato no mesmo dia, depois das 12h no dia útil seguinte. No contato a Saúde confere as vacinas, agenda o exame (rede credenciada, clínica próxima ou reembolso) e dá o parecer. Prazo do ASO: 5 dias úteis só com exame clínico, 10 com exames complementares (pela tabela de exames da função). Sem contato ou falta por 24h, recusa de vacina, segunda remarcação ou erro no SOC vão para o R&S decidir. Clique na pessoa para registrar. O ASO oficial continua no SOC.`}</p></div></div>
 <div class="caixa pilha"><h2>Fila de agendamento</h2>${filaSaude()}</div>
 ${vacinasSaude()}
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
    <a class="bt" href="${(window.ESTEIRA_URL || 'esteira.html') + '?v=38' + modoEsteira()}" target="_blank" rel="noopener">Abrir em tela cheia</a></div>
  ${ondeEstaCadaKit('todas')}`;
  return `<div class="cab"><div class="t"><span class="olho">Kit admissional</span><h1>Esteira de assinatura</h1><p class="apoio">A mesma Esteira de Admissão de sempre: junta os PDFs do LG por pessoa, a Cimed assina primeiro e o colaborador assina na tela. Com <b>Publicar on-line</b>, o kit vai para o armazenamento privado da plataforma: a Cimed assina de qualquer aparelho e o colaborador assina pelo QR, no celular dele, depois de confirmar CPF e data de nascimento. Sem publicar, o kit fica só neste aparelho, como antes.</p></div>
    <a class="bt" href="${(window.ESTEIRA_URL || 'esteira.html') + '?v=38' + modoEsteira()}" target="_blank" rel="noopener">Abrir em tela cheia</a></div>
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
    host.innerHTML = `<iframe class="esteira" src="${(window.ESTEIRA_URL || 'esteira.html') + '?v=38' + modoEsteira()}" title="Esteira de Admissão"></iframe>`;
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
  const doDia = vivas().filter((a) => a.data_inicio === E.turmaLista && !a.teste);
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
  ${cal.length ? cal.map((t) => `<div class="faixa"><p><b>${esc(t.nome || GRUPO_TURMA[t.grupo])}</b> · inclusão até ${dmh(t.corte_inclusao)}${t.corte_remuneracao ? ' · Remuneração até ' + dmh(t.corte_remuneracao) : ''} · documentos, ASO e LG até ${dmh(t.corte_documentos)}${t.formato ? ' · ' + (t.formato === 'online' ? 'on-line' : 'presencial') : ''}${t.link_ou_endereco ? ' · ' + esc(t.link_ou_endereco) : ''}${t.contato_chegada ? ' · procurar ' + esc(t.contato_chegada) : ''}${t.concluida_avisada_em ? ' · ' + chip('c-ok', 'ok', 'áreas avisadas: admissões concluídas ' + dm(new Date(t.concluida_avisada_em))) : ''}</p>${tem('admissao') && temSql27() ? `<button class="bt peq" data-act="turmaConcluida" data-id="${t.id}" title="E-mail a todas as áreas com quem foi admitido nesta turma">${t.concluida_avisada_em ? 'Comunicar de novo' : 'Comunicar admissões concluídas'}</button>` : ''}</div>`).join('') : (E.turmaLista ? '<div class="aviso-faixa">Esta data não está no cronograma de turmas. Cadastre na tela "Cronograma e corte" para o corte e o convite usarem os horários certos.</div>' : '')}
  <div class="caixa pilha"><div class="linha-flex"><h2 style="flex:1">${plural(lista.length, 'pessoa', 'pessoas')} na lista${semMat && E.soMatriculados ? ` <span class="apoio" style="font-size:13px">· ${plural(semMat, 'ainda sem matrícula', 'ainda sem matrícula')}</span>` : ''}</h2>
    <button class="bt" data-act="listaXlsx" ${lista.length ? '' : 'disabled'}>Baixar .xlsx</button><button class="bt" data-act="listaCsv" ${lista.length ? '' : 'disabled'}>Baixar .csv</button>
    ${tem('admissao') ? `<button class="bt pri" data-act="listaEnviar" ${lista.length && E.soMatriculados ? '' : 'disabled'} title="${E.soMatriculados ? '' : 'Marque só matriculados para enviar'}">Enviar por e-mail</button>` : ''}</div>
    ${tem('admissao') ? `<p class="apoio" style="font-size:12.5px">Vai para: ${destinos.length ? esc(destinos.join(', ')) : '<b>ninguém cadastrado para estas unidades</b>'}${reqs.length ? ' · em cópia os requisitantes: ' + esc(reqs.join(', ')) : ''}. Cadastre e ajuste em "Especialistas e avisos".</p>` : ''}
    ${lista.length ? `<div class="tab-wrap"><table class="t-lista"><thead><tr>${COLS_LISTA.map((c) => `<th>${esc(c[1])}</th>`).join('')}</tr></thead><tbody>${lista.map((x) => `<tr class="clic" data-act="ficha" data-id="${x.id}">${COLS_LISTA.map(([k]) => `<td>${esc(k === 'data_admissao' ? dm(x[k]) : x[k] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Ninguém para listar com estes filtros.</div>'}</div>
  `;
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
  return `<div class="caixa pilha" id="cronograma"><h2>Cronograma de turmas</h2><p class="apoio" style="font-size:13px">O cronograma mensal do DP, dentro da plataforma: o cadastro escolhe a turma, o contador de corte e as cobranças usam estes horários, e o convite usa o formato, o link ou endereço e quem procurar. Regra sugerida (cronograma do mês): inclusão até 5 dias úteis antes, às 12h; a Remuneração valida até o fim do mesmo dia; documentos, ASO e LG até 4 dias antes, ao meio-dia. Semana com feriado: ajuste à mão.</p>
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
  // cronograma do Time de Admissão: inclusão 5 dias úteis antes, às 12h; a Remuneração valida até o fim do mesmo dia (feriado: ajuste à mão)
  const uteis = (k, h, m = 0) => { const x = new Date(base); let n = k; while (n > 0) { x.setDate(x.getDate() - 1); if (x.getDay() % 6) n--; } x.setHours(h, m, 0, 0); return paraLocal(x.toISOString()); };
  if (!$('#tnInc').value) $('#tnInc').value = uteis(5, 12);
  if (!$('#tnRem').value) $('#tnRem').value = uteis(5, 23, 59);
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

// --------------------------------------------------------- cronograma e corte
// Substitui o e-mail "Cronograma admissões" do mês: a recrutadora vê as datas aqui e no cadastro.
VIEWS.cronograma = () => {
  const hoje = iso(new Date());
  const ts = E.turmasCal.filter((t) => t.ativo && t.data >= hoje).sort((x, y) => x.data.localeCompare(y.data) || x.grupo.localeCompare(y.grupo));
  const leitura = `<div class="caixa pilha"><h2>Próximas integrações</h2>${ts.length ? `<div class="tab-wrap"><table><thead><tr><th>Data de admissão</th><th>Unidades</th><th>Inclusão (corte)</th><th>Remuneração valida até</th><th>Documentos, ASO e LG até</th><th>Onde</th></tr></thead><tbody>${ts.map((t) => { const tarde = Date.now() > new Date(t.corte_inclusao).getTime(); return `<tr><td class="num"><b>${dmd(t.data)}</b></td><td>${esc(t.nome || GRUPO_TURMA[t.grupo])}</td><td>${tarde ? chip('c-ag', '', 'encerrado · ' + dmh(t.corte_inclusao)) : chip('c-at', 'at', 'até ' + dmh(t.corte_inclusao))}</td><td class="num">${t.corte_remuneracao ? dmh(t.corte_remuneracao) : '—'}</td><td class="num">${dmh(t.corte_documentos)}</td><td style="font-size:12.5px">${t.formato ? (t.formato === 'online' ? 'on-line' : 'presencial') : 'padrão do local'}${t.link_ou_endereco ? ' · ' + esc(t.link_ou_endereco) : ''}</td></tr>`; }).join('')}</tbody></table></div>` : '<div class="vazio">O Time de Admissão ainda não cadastrou as próximas turmas.</div>'}</div>`;
  const minhas = (E.excecoes || []).filter((x) => !x.cancelada_em).slice(0, 60);
  return `<div class="cab"><div class="t"><span class="olho">Admissão</span><h1>Cronograma e corte</h1><p class="apoio">As datas de integração e de corte de cada turma, cadastradas pelo Time de Admissão. Substitui o e-mail "Cronograma admissões" do mês: a recrutadora consulta aqui, e o cadastro mostra o corte da data escolhida. ${temSql27() ? 'Depois do corte o cadastro trava; exceção aprovada, só o Time de Admissão libera, por CPF.' : ''}</p></div></div>
  ${tem('admissao') ? cronogramaTurmas() : leitura}
  ${temSql27() && tem('admissao') ? excecoesBox() : ''}
  ${temSql27() && !tem('admissao') && minhas.length ? `<div class="caixa pilha"><h2>Exceções liberadas</h2><div class="tab-wrap"><table><thead><tr><th>CPF</th><th>Nome</th><th>Turma</th><th>Situação</th></tr></thead><tbody>${minhas.map((x) => `<tr><td class="num">${fmtCpf(x.cpf)}</td><td>${esc(titulo(x.nome || ''))}</td><td class="num">${dmd(x.data_inicio)}</td><td>${x.usada_em ? chip('c-ok', 'ok', 'usada ' + dm(new Date(x.usada_em))) : chip('c-at', 'at', 'liberada: pode cadastrar')}</td></tr>`).join('')}</tbody></table></div></div>` : ''}`;
};
function excecoesBox() {
  const hoje = iso(new Date());
  const ts = E.turmasCal.filter((t) => t.ativo && t.data >= hoje).sort((x, y) => x.data.localeCompare(y.data));
  const datas = [...new Set(ts.map((t) => t.data))];
  const l = (E.excecoes || []).slice(0, 80);
  return `<div class="caixa pilha" id="excecoes"><h2>Exceções ao corte</h2><p class="apoio" style="font-size:13px">Inclusão aprovada depois do corte: informe o CPF, a turma e o motivo (quem aprovou e por quê). A recrutadora passa a ver a turma liberada para esse CPF no cadastro. Cada liberação vale uma vez.</p>
    <form id="fExc" class="form" style="grid-template-columns:repeat(4,minmax(0,1fr))">
      <div class="campo"><label for="exCpf">CPF</label><input id="exCpf" inputmode="numeric" placeholder="000.000.000-00" required></div>
      <div class="campo"><label for="exNome">Nome</label><input id="exNome" placeholder="Opcional"></div>
      <div class="campo"><label for="exData">Turma (data de admissão)</label><select id="exData">${datas.map((d) => `<option value="${d}">${dmd(d)}</option>`).join('')}<option value="outra">Outra data…</option></select><input id="exOutra" type="date" hidden style="margin-top:6px"></div>
      <div class="campo"><label for="exMotivo">Motivo</label><input id="exMotivo" placeholder="Ex.: aprovado pela gerência em 21/10" required></div>
      <div class="campo l4" style="align-items:flex-end"><button class="bt pri" type="submit">Liberar inclusão</button></div>
    </form>
    ${l.length ? `<div class="tab-wrap"><table><thead><tr><th>CPF</th><th>Nome</th><th>Turma</th><th>Motivo</th><th>Liberado por</th><th>Situação</th><th></th></tr></thead><tbody>${l.map((x) => `<tr ${x.cancelada_em ? 'style="opacity:.55"' : ''}><td class="num">${fmtCpf(x.cpf)}</td><td>${esc(titulo(x.nome || ''))}</td><td class="num">${dmd(x.data_inicio)}</td><td style="font-size:12.5px">${esc(x.motivo)}</td><td style="font-size:12.5px">${esc(x.liberado_por)} · ${quando(x.criado_em)}</td><td>${x.cancelada_em ? chip('c-ag', '', 'cancelada') : x.usada_em ? chip('c-ok', 'ok', 'usada por ' + (x.usada_por || '') + ' ' + dm(new Date(x.usada_em))) : chip('c-at', 'at', 'aguardando o cadastro')}</td><td>${!x.usada_em && !x.cancelada_em ? `<button class="bt peq fan" data-act="excCancelar" data-id="${x.id}">Cancelar</button>` : ''}</td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nenhuma exceção liberada.</div>'}</div>`;
}
POS.cronograma = () => {
  $('#fTurmaNova')?.addEventListener('submit', salvarTurmaForm);
  $('#tnData')?.addEventListener('change', sugerirCortes); $('#tnGrupo')?.addEventListener('change', sugerirCortes);
  $('#exData')?.addEventListener('change', (e) => { $('#exOutra').hidden = e.target.value !== 'outra'; });
  $('#fExc')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const cpf = $('#exCpf').value.replace(/\D/g, ''); const data = $('#exData').value === 'outra' ? $('#exOutra').value : $('#exData').value;
    if (!cpfValido(cpf)) { toast('CPF inválido.', true); return; }
    if (!data) { toast('Escolha a turma.', true); return; }
    if (await tentar(() => api.excecaoLiberar(cpf, $('#exNome').value.trim(), data, $('#exMotivo').value.trim()), 'Inclusão liberada para este CPF na turma de ' + dm(data))) await recarregar();
  });
};

// --------------------------------------------------------- listas e regras (Admissão)
const TIPOS_LISTA = [['analista_rs', 'Analistas de R&S'], ['escala', 'Escalas'], ['turno', 'Turnos'], ['fretado', 'Linhas do fretado'], ['conferente_docs', 'Conferência dos documentos'], ['agendador_aso', 'Agendamento do exame'], ['aso_obs', 'Observações do ASO'], ['sindicato', 'Sindicatos'], ['acesso_ti', 'Equipamentos e acessos']];
VIEWS.listas = () => {
  const tipo = E.listaTipo; const itens = (E.listas || []).filter((x) => x.tipo === tipo);
  const fabricas = E.locais.filter((l) => l.tipo === 'fabrica').map((l) => l.nome);
  const linha = (x) => `<div class="area" data-lid="${x.id || ''}"><div><b>${x.id ? esc(x.valor) : 'Novo item'}</b>${x.filtro ? `<div class="apoio" style="font-size:12px">${esc(x.filtro)}</div>` : ''}</div>
    <div class="ac" style="grid-column:span 2"><input class="sel-doc" data-f="valor" value="${esc(x.valor || '')}" style="width:100%" placeholder="Texto que aparece na lista">${tipo === 'fretado' ? `<select class="sel-doc" data-f="filtro" style="margin-top:4px">${fabricas.map((f) => `<option ${f === x.filtro ? 'selected' : ''}>${esc(f)}</option>`).join('')}</select>` : ''}</div>
    <div class="linha-flex"><label class="marcas"><label><input type="checkbox" data-f="ativo" ${x.ativo !== false ? 'checked' : ''}> Ativo</label></label><button class="bt peq" data-act="listaSalvar">Salvar</button>${x.id ? `<button class="bt peq fan" data-act="listaApagar" data-id="${x.id}">Apagar</button>` : ''}</div></div>`;
  return `<div class="cab"><div class="t"><span class="olho">Admissão</span><h1>Listas e regras</h1><p class="apoio">As listas suspensas do cadastro e da ficha, mantidas pelo Time de Admissão (vieram da aba ANALISTAS da planilha). Item inativo some da lista, mas continua nas fichas antigas. As regras de sindicato por local ensinam o sistema a conferir o que a recrutadora escolheu.</p></div></div>
  <div class="seg" role="group" aria-label="Lista" style="flex-wrap:wrap;margin-bottom:12px">${TIPOS_LISTA.map(([k, n]) => `<button data-act="listaTipo" data-t="${k}" aria-pressed="${k === tipo}">${esc(n)} <span class="num">${(E.listas || []).filter((x) => x.tipo === k).length}</span></button>`).join('')}</div>
  <div class="caixa" style="padding:0"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>${esc((TIPOS_LISTA.find((x) => x[0] === tipo) || [])[1] || '')}</h2></div>${itens.map(linha).join('')}${linha({ tipo, filtro: tipo === 'fretado' ? fabricas[0] : null })}</div>
  ${regrasSindicatoBox()}`;
};
function regrasSindicatoBox() {
  const linha = (r) => `<div class="area" data-rid="${r.id || ''}"><div><b>${r.id ? esc(r.local) : 'Nova regra'}</b>${r.cargo_contem ? `<div class="apoio" style="font-size:12px">cargo contém "${esc(r.cargo_contem)}"</div>` : ''}</div>
    <div class="ac" style="grid-column:span 2"><div class="linha-flex" style="gap:6px"><select class="sel-doc" data-f="local">${E.locais.map((l) => `<option ${l.nome === r.local ? 'selected' : ''}>${esc(l.nome)}</option>`).join('')}</select><input class="sel-doc" data-f="cargo_contem" value="${esc(r.cargo_contem || '')}" placeholder="cargo contém (opcional)" style="width:170px"></div><input class="sel-doc" data-f="sindicato" value="${esc(r.sindicato || '')}" placeholder="Sindicato" style="width:100%;margin-top:4px"><input class="sel-doc" data-f="observacao" value="${esc(r.observacao || '')}" placeholder="Observação (convenção, CNPJ…)" style="width:100%;margin-top:4px"></div>
    <div class="linha-flex"><button class="bt peq" data-act="regraSalvar">Salvar</button>${r.id ? `<button class="bt peq fan" data-act="regraApagar" data-id="${r.id}">Apagar</button>` : ''}</div></div>`;
  return `<div class="caixa" style="padding:0" id="regrasSind"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>Regras de sindicato por local</h2><p class="apoio" style="font-size:12.5px">Uma linha por local (e, se preciso, por cargo). O cadastro compara o sindicato escolhido com a regra: igual fica "conforme"; diferente ou sem regra vai para Relações Sindicais conferir. Regra a confirmar com o Ico: base territorial do local de trabalho e empresa; comercial externo segue a cidade onde atua.</p></div>${(E.regrasSind || []).map(linha).join('')}${linha({ local: E.locais[0]?.nome })}</div>`;
}

// --------------------------------------------------------- sindicatos (Relações Sindicais)
VIEWS.sindicatos = () => {
  const l = E.admissoes.filter((a) => !a.declinou_em && !a.matricula && (E.sindTodos || ['divergente', 'sem_regra', 'pendente'].includes(a.sindicato_status))).sort((x, y) => (x.data_inicio + x.local).localeCompare(y.data_inicio + y.local));
  const sinds = [...new Set([...opcoesLista('sindicato'), ...(E.regrasSind || []).map((r) => r.sindicato)])].sort();
  return `<div class="cab"><div class="t"><span class="olho">Relações Sindicais</span><h1>Sindicatos</h1><p class="apoio">O sindicato que a recrutadora escolheu, comparado com a regra do local. Confira o que está diferente da regra ou sem regra: você pode corrigir antes de conferir. A conferência fica com seu nome e a data na ficha.</p></div></div>
  <div class="filtros"><label class="marcas"><label><input type="checkbox" id="sindTodos" ${E.sindTodos ? 'checked' : ''}> Mostrar também os conformes e os já conferidos</label></label></div>
  <datalist id="dlSindL">${sinds.map((s) => `<option value="${esc(s)}">`).join('')}</datalist>
  <div class="caixa pilha">${l.length ? `<div class="tab-wrap"><table><thead><tr><th>Pessoa</th><th>Cargo · local</th><th>Integração</th><th>Escolhido</th><th>Regra do local</th><th>Situação</th><th>Conferir</th></tr></thead><tbody>${l.map((a) => `<tr><td class="clic" data-act="ficha" data-id="${a.id}"><b>${esc(titulo(a.nome))}</b><br><span class="apoio" style="font-size:12px">${esc(a.analista_rs || '')}</span></td><td>${esc(a.cargo)}<br><span class="apoio" style="font-size:12px">${esc(a.local)}${a.setor ? ' · ' + esc(a.setor) : ''}</span></td><td class="num">${dmd(a.data_inicio)}</td><td>${esc(a.sindicato || '—')}</td><td>${esc(a.sindicato_sugerido || '—')}</td><td>${chipSind(a)}${a.sindicato_conferido_por ? `<br><span class="apoio" style="font-size:11.5px">${esc(a.sindicato_conferido_por)}</span>` : ''}</td><td><div class="linha-flex" style="gap:4px"><input class="sel-doc" list="dlSindL" data-sind="${a.id}" value="${esc(a.sindicato || a.sindicato_sugerido || '')}" style="width:190px" aria-label="Sindicato"><button class="bt peq" data-act="sindConferirLinha" data-id="${a.id}">Conferir</button></div></td></tr>`).join('')}</tbody></table></div>` : '<div class="vazio">Nada para conferir agora.</div>'}</div>
  ${regrasSindicatoBox()}`;
};
POS.sindicatos = () => { $('#sindTodos')?.addEventListener('change', (e) => { E.sindTodos = e.target.checked; render(); }); };

// --------------------------------------------------------- Benefícios (SQL 28)
const BEN_ST = { previsto: ['c-ag', '', 'Previsto'], aguardando_escolha: ['c-ag', 'ag', 'Aguardando escolha'], escolhido: ['c-info', 'info', 'Escolhas recebidas'], recusado: ['c-ag', '', 'Recusou'], pendente_documento: ['c-at', 'at', 'Falta documento'], pronto: ['c-info', '', 'Pronto para envio'], enviado: ['c-info', 'ag', 'Enviado'], ativo: ['c-ok', 'ok', 'Ativo'], com_pendencia: ['c-cr', 'cr', 'Com pendência'], cancelado: ['c-ag', 'cr', 'Cancelado'] };
const BEN_ORDEM = ['com_pendencia', 'pendente_documento', 'aguardando_escolha', 'escolhido', 'pronto', 'previsto', 'enviado', 'ativo', 'recusado', 'cancelado'];
const FAMILIAS_BEN = [['saude', 'Saúde'], ['odonto', 'Odonto'], ['seguro', 'Seguro'], ['transporte', 'Transporte'], ['alimentacao', 'VA · VR · restaurante'], ['opcionais', 'Opcionais']];
const benDe = (id) => (E.benPessoa || []).filter((x) => x.admissao_id === id);
const benCat = (c) => (E.benCat || []).find((x) => x.codigo === c) || { nome: c, familia: 'opcionais' };
const chipBen = (st, txt) => { const s = BEN_ST[st] || BEN_ST.previsto; return chip(s[0], s[1], txt || s[2]); };
function situacaoBen(its) {
  if (!its.length) return ['c-ag', '', 'Sem pacote (temporário ou a definir)'];
  if (its.some((x) => ['com_pendencia', 'pendente_documento'].includes(x.status))) return ['c-cr', 'cr', 'Com pendência'];
  if (its.every((x) => x.status === 'previsto')) return ['c-ag', '', 'Previsão (sem matrícula)'];
  if (its.every((x) => ['ativo', 'recusado', 'cancelado'].includes(x.status))) return ['c-ok', 'ok', 'Concluído'];
  if (its.some((x) => x.status === 'aguardando_escolha')) return ['c-at', 'at', 'Aguardando escolhas'];
  return ['c-info', 'info', 'Em inclusão'];
}
const depsDe = (id) => (E.depsTodos || []).filter((d) => d.admissao_id === id);
const opcDe = (id) => (E.opcTodas || []).find((o) => o.admissao_id === id);
VIEWS.beneficios = () => {
  const comItens = new Set((E.benPessoa || []).map((x) => x.admissao_id));
  const datas = [...new Set(E.admissoes.filter((a) => comItens.has(a.id)).map((a) => a.data_inicio))].sort();
  if (!E.benTurma || (E.benTurma !== 'todas' && !datas.includes(E.benTurma))) E.benTurma = datas.find((d) => d >= iso(new Date())) || datas[datas.length - 1] || 'todas';
  const ps = E.admissoes.filter((a) => comItens.has(a.id) && (E.benTurma === 'todas' || a.data_inicio === E.benTurma)).sort((x, y) => (x.local + x.nome).localeCompare(y.local + y.nome));
  const its = (E.benPessoa || []).filter((x) => ps.some((a) => a.id === x.admissao_id));
  const k = (f) => ps.filter(f).length;
  const saudeEnviar = its.filter((x) => benCat(x.beneficio).familia === 'saude' && ['escolhido', 'pronto'].includes(x.status)).length;
  const fretadoAtivar = its.filter((x) => x.beneficio === 'fretado' && ['escolhido', 'pronto', 'enviado'].includes(x.status)).length;
  const cellFam = (a, fam) => { const l = benDe(a.id).filter((x) => benCat(x.beneficio).familia === fam); if (!l.length) return '<span class="apoio">—</span>'; const pior = l.slice().sort((x, y) => BEN_ORDEM.indexOf(x.status) - BEN_ORDEM.indexOf(y.status))[0]; return `<span title="${esc(l.map((x) => benCat(x.beneficio).nome + ': ' + (BEN_ST[x.status] || [])[2]).join('\n'))}">${chipBen(pior.status)}</span>`; };
  return `<div class="cab"><div class="t"><span class="olho">Benefícios</span><h1>Benefícios da turma</h1><p class="apoio">O pacote previsto de cada pessoa pelo grupo do Painel de Benefícios (fábrica, escritórios SP, comercial, CD, aprendiz, estagiário), desde o cadastro. Com a matrícula, os itens esperam a escolha do kit; quando o kit é assinado, as escolhas chegam aqui. Marque cada item até ficar ativo. Valores e regras finas continuam no Painel; o LG segue sendo o registro oficial.</p></div></div>
  <div class="filtros"><label class="olho" for="bTurma">Turma</label><select id="bTurma"><option value="todas" ${E.benTurma === 'todas' ? 'selected' : ''}>Todas</option>${datas.map((d) => `<option value="${d}" ${d === E.benTurma ? 'selected' : ''}>${dmd(d)}</option>`).join('')}</select>
    <span style="flex:1"></span><button class="bt peq" data-act="benCsv" data-k="saude" title="O LG gera o WKF; esta planilha confere quem entra e o que falta">Conferência para o WKF · saúde e odonto (.csv)</button><button class="bt peq" data-act="benCsv" data-k="fretado">Fretado · Rota (.csv)</button><button class="bt peq" data-act="benCsv" data-k="vt">Vale-transporte (.csv)</button><button class="bt peq" data-act="benCsv" data-k="tudo">Checklist completo (.csv)</button></div>
  <div class="grade g4">
    <div class="caixa kpi"><b class="num">${ps.length}</b><span>pessoas na seleção</span><small class="apoio" style="font-size:12px">${k((a) => a.matricula)} com matrícula</small></div>
    <div class="caixa kpi"><b class="num">${k((a) => opcDe(a.id))}</b><span>escolhas do kit recebidas</span><small class="apoio" style="font-size:12px">${k((a) => depsDe(a.id).length)} com dependentes</small></div>
    <div class="caixa kpi${saudeEnviar ? ' alerta' : ''}"><b class="num">${saudeEnviar}</b><span>inclusões de saúde a enviar</span><small class="apoio" style="font-size:12px">Unimed sobe D+1 da integração</small></div>
    <div class="caixa kpi"><b class="num">${k((a) => situacaoBen(benDe(a.id))[2] === 'Com pendência')}</b><span>com pendência</span><small class="apoio" style="font-size:12px">${fretadoAtivar} fretado(s) sem QR Code confirmado</small></div>
  </div>
  <div class="caixa pilha"><div class="linha-flex" style="gap:6px"><b style="flex:1">Marcar vários de uma vez</b><select id="bLoteBen" class="sel-doc">${(E.benCat || []).map((c) => `<option value="${esc(c.codigo)}">${esc(c.nome)}</option>`).join('')}</select><select id="bLoteSt" class="sel-doc">${['enviado', 'ativo', 'pronto', 'recusado'].map((s) => `<option value="${s}">${BEN_ST[s][2]}</option>`).join('')}</select><button class="bt peq pri" data-act="benLote">Aplicar aos marcados</button></div>
   ${ps.length ? `<div class="tab-wrap"><table><thead><tr><th><input type="checkbox" id="bTodos" aria-label="Marcar todos"></th><th>Pessoa</th><th>Grupo</th>${FAMILIAS_BEN.map(([, n]) => `<th>${esc(n)}</th>`).join('')}<th>Dependentes</th><th>Situação</th></tr></thead><tbody>${ps.map((a) => { const s = situacaoBen(benDe(a.id)); const d = depsDe(a.id); return `<tr><td>${a.declinou_em ? '' : `<input type="checkbox" data-bsel="${a.id}" aria-label="Marcar ${esc(titulo(a.nome))}">`}</td><td class="clic" data-act="benAbrir" data-id="${a.id}"><b>${esc(titulo(a.nome))}</b>${a.declinou_em ? ' ' + chip('c-cr', 'cr', 'cancelada') : ''}<br><span class="apoio" style="font-size:12px">${esc(a.cargo)} · ${esc(a.local)}${a.matricula ? ' · ' + esc(a.matricula) : ''}</span></td><td style="font-size:12px">${esc(grupoBen(a))}</td>${FAMILIAS_BEN.map(([f]) => `<td>${cellFam(a, f)}</td>`).join('')}<td class="num">${d.length || '—'}</td><td>${chip(s[0], s[1], s[2])}</td></tr>`; }).join('')}</tbody></table></div>` : '<div class="vazio">Ninguém com pacote de benefícios nesta seleção.</div>'}</div>
  ${painelFretado()}`;
};
// grupo do Painel (o mesmo do banco, beneficios_grupo)
function grupoBen(a) {
  if (a.tipo_contrato === 'Jovem aprendiz' && localDe(a.local).tipo === 'fabrica') return a.local === 'Fernão Dias' ? 'Aprendiz Unidade 2 (fretado)' : 'Aprendiz fábrica (VT)';
  const tp = localDe(a.local).tipo;
  if (a.tipo_contrato === 'Temporário') return 'Temporário (a definir)';
  if (a.tipo_contrato === 'Jovem aprendiz') return tp === 'fabrica' ? 'Aprendiz Pouso' : tp === 'cd' && a.local !== 'CD SP' ? 'Aprendiz CD' : 'Aprendiz SP';
  if (a.tipo_contrato === 'Estágio') return tp === 'fabrica' ? 'Estagiário Pouso' : 'Estagiário escritórios';
  return ({ fabrica: 'Pouso · SSBV', sp: 'Escritórios SP', comercial: 'Comercial', cd: 'CD ' + String(a.local || '').replace('CD ', '') })[tp] || 'A definir';
}
// fretado: as linhas do LG e os itinerários por unidade e turno (SQL 31), para consulta
function painelFretado() {
  if (!temSql31()) return '';
  if (!E.itins) { if (!E.itinsP) garantirItins().then(() => { if (E.view === 'beneficios') render(); }); return '<div class="caixa pilha"><h2>Fretado · linhas e itinerários</h2><span class="apoio">Carregando…</span></div>'; }
  const unis = ['UNIDADE 1', 'UNIDADE 2', 'SSBV']; const turnos = [...new Set(E.itins.filter((i) => i.unidade === E.itinUni).map((i) => i.turno))].sort();
  if (!turnos.includes(E.itinTurno)) E.itinTurno = turnos[0] || '';
  const l = E.itins.filter((i) => i.unidade === E.itinUni && i.turno === E.itinTurno);
  const semItin = linhasFretado(E.itinUni === 'UNIDADE 1' ? 'Pouso Alegre' : E.itinUni === 'UNIDADE 2' ? 'Fernão Dias' : 'Gráfica SSBV').filter((x) => !E.itins.some((i) => i.codigo_lg === x.codigo_lg));
  return `<div class="caixa pilha"><h2>Fretado · linhas e itinerários</h2><p class="apoio" style="font-size:12.5px">Linhas do LG (24 - FRETADO) e os itinerários de ida por unidade e turno, enviados por Benefícios em 21/10/2026. Unidade 1 = Pouso Alegre (Site I), Unidade 2 = Fernão Dias (Site II). Pontos técnicos de correção de trajeto não aparecem.</p>
    <div class="filtros"><label class="olho" for="itUni">Unidade</label><select id="itUni">${unis.map((u) => `<option ${u === E.itinUni ? 'selected' : ''}>${u}</option>`).join('')}</select><label class="olho" for="itTurno">Turno</label><select id="itTurno">${turnos.map((x) => `<option ${x === E.itinTurno ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select><button class="bt peq" data-act="benItinFiltro">Ver</button></div>
    ${l.map(tabelaItin).join('') || '<div class="vazio">Sem itinerário para este turno.</div>'}
    ${semItin.length ? `<p class="apoio" style="font-size:12px">Linhas do LG desta unidade sem itinerário enviado: ${semItin.map((x) => esc(nomeLinha(x)) + ' (LG ' + x.codigo_lg + ')').join(', ')}.</p>` : ''}</div>`;
}
POS.beneficios = () => {
  $('#bTurma')?.addEventListener('change', (e) => { E.benTurma = e.target.value; render(); });
  $('#bTodos')?.addEventListener('change', (e) => document.querySelectorAll('[data-bsel]').forEach((c) => { c.checked = e.target.checked; }));
};
async function abrirBen(id) {
  const a = adm(id); if (!a) return;
  if ((a.fretado || a.fretado_codigo || a.fretado_itin_chave || a.fretado_integ) && UNID_FRETADO[a.local]) await garantirItins();
  const its = benDe(id).slice().sort((x, y) => (benCat(x.beneficio).ordem || 0) - (benCat(y.beneficio).ordem || 0));
  const deps = depsDe(id); const opc = opcDe(id);
  const alerta = (d) => (d.parentesco === 'Filho(a)' || d.parentesco === 'Enteado(a)') && d.idade != null ? (d.idade >= 24 ? 'acima de 24: não entra no plano' : d.idade >= 21 && !d.universitario ? 'de 21 a 24: só universitário com declaração' : d.idade >= 21 ? 'universitário: declaração semestral' : '') : (['Pai', 'Mãe'].includes(d.parentesco) ? 'pais: só no odonto SulAmérica' : '');
  $('#camada').innerHTML = `<div class="veu" data-act="fechar"><aside class="gaveta" data-pare="1"><div class="pilha">
   <div class="linha-flex" style="align-items:flex-start"><div class="pilha" style="flex:1;gap:6px"><span class="olho">Benefícios · ${esc(grupoBen(a))}</span><h2 style="font-size:22px">${esc(titulo(a.nome))}</h2><span class="apoio">${esc(a.cargo)} · ${esc(a.local)} · integração ${dmd(a.data_inicio)}${a.matricula ? ' · matrícula ' + esc(a.matricula) : ' · sem matrícula ainda'}${a.cpf ? ' · CPF ' + fmtCpf(a.cpf) : ''}</span><span class="apoio">${esc(a.estado_civil || '')}${a.fretado ? ' · fretado: ' + esc(a.fretado) + (a.fretado_codigo ? ' (LG ' + a.fretado_codigo + ')' : '') : ''}${a.turno ? ' · ' + esc(a.turno) : ''}</span>${a.era_dependente_cimed ? `<div class="motivo">Já era dependente de alguém da Cimed no plano: mover de dependente para titular no WKF da operadora.</div>` : ''}</div><button class="bt" data-act="fechar">Fechar</button></div>
   <div class="caixa pilha"><h3>Pacote e checklist</h3><div class="docs">${its.map((x) => { const c = benCat(x.beneficio); return `<div class="doc"><div><div class="dn">${esc(c.nome)}${c.natureza === 'compulsorio' ? ' <span class="apoio" style="font-weight:400">· compulsório</span>' : c.natureza === 'regra' ? ' <span class="apoio" style="font-weight:400">· pela regra do grupo</span>' : ''}</div><div class="dm">${chipBen(x.status)}${x.plano ? ' · ' + esc(x.plano) : ''} <span class="apoio" style="font-size:12px">${esc(c.vigencia || '')}</span>${x.wkf_ref ? `<div class="motivo">WKF: ${esc(x.wkf_ref)}${x.enviado_em ? ' · enviado ' + quando(x.enviado_em) : ''}</div>` : ''}${x.pendencia ? `<div class="motivo">${esc(x.pendencia)}</div>` : ''}${x.atualizado_por && x.atualizado_por !== 'Sistema' ? `<div class="motivo">${esc(x.atualizado_por)} · ${quando(x.atualizado_em)}</div>` : ''}</div></div>
     <div class="acoes" data-benbox="${esc(x.beneficio)}"><select class="sel-doc" data-f="st" aria-label="Situação">${x.status === 'previsto' ? '<option value="previsto" selected disabled>Previsto (sem matrícula)</option>' : ''}${['aguardando_escolha', 'escolhido', 'recusado', 'pendente_documento', 'pronto', 'enviado', 'ativo', 'com_pendencia', 'cancelado'].map((s) => `<option value="${s}" ${s === x.status ? 'selected' : ''}>${BEN_ST[s][2]}</option>`).join('')}</select><input class="sel-doc" data-f="plano" value="${esc(x.plano || '')}" placeholder="plano" style="width:120px" aria-label="Plano"><input class="sel-doc" data-f="pend" value="${esc(x.pendencia || '')}" placeholder="pendência" style="width:140px" aria-label="Pendência"><button class="bt peq" data-act="benSalvar" data-id="${a.id}" data-b="${esc(x.beneficio)}">Salvar</button></div></div>`; }).join('') || '<div class="vazio">Sem pacote previsto (temporário ou local a definir).</div>'}</div></div>
   ${(a.fretado || a.fretado_codigo || a.fretado_itin_chave || a.fretado_integ) && UNID_FRETADO[a.local] && temSql31() ? `<div class="caixa pilha" style="gap:6px"><h3>Fretado · itinerários</h3>${blocoItinerario(a)}</div>` : ''}
   <div class="caixa pilha"><h3>Dependentes · ${deps.length}</h3>${deps.length ? `<div class="tab-wrap"><table><thead><tr><th>Nome</th><th>Parentesco</th><th>Idade</th><th>CPF</th><th>Plano</th><th>Odonto</th><th>Regra</th></tr></thead><tbody>${deps.map((d) => `<tr><td>${esc(titulo(d.nome))}</td><td>${esc(d.parentesco)}</td><td class="num">${d.idade ?? '—'}</td><td class="num">${d.cpf ? fmtCpf(d.cpf) : '—'}</td><td>${d.plano_saude ? 'sim' : '—'}</td><td>${d.odonto ? 'sim' : '—'}</td><td style="font-size:12px">${esc(alerta(d))}</td></tr>`).join('')}</tbody></table></div>` : '<span class="apoio">Nenhum dependente informado no cadastro.</span>'}</div>
   <div class="caixa pilha"><h3>Escolhas feitas no kit</h3>${opc ? `<span class="apoio" style="font-size:12.5px">Recebidas em ${quando(opc.recebido_em)}</span><div class="grade g2" style="gap:4px 16px;font-size:13px">${resumoOpcoes(opc.dados) || '<span class="apoio">nenhuma escolha registrada</span>'}</div>` : '<span class="apoio">Chegam aqui quando o colaborador assina o kit.</span>'}</div>
   ${tem('rs', 'admissao') ? `<div><button class="bt" data-act="ficha" data-id="${a.id}">Abrir a ficha completa</button></div>` : ''}
  </div></aside></div>`;
}
function baixarCsvSimples(nome, cab, linhas) {
  const q = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  baixarBlob(nome, new Blob(['﻿' + [cab, ...linhas].map((l) => l.map(q).join(';')).join('\r\n')], { type: 'text/csv;charset=utf-8' }));
}
function exportarBen(k) {
  const comItens = new Set((E.benPessoa || []).map((x) => x.admissao_id));
  const ps = E.admissoes.filter((a) => !a.declinou_em && comItens.has(a.id) && (E.benTurma === 'todas' || a.data_inicio === E.benTurma));
  const sufixo = E.benTurma === 'todas' ? 'todas' : E.benTurma;
  if (k === 'saude') {
    const linhas = [];
    ps.forEach((a) => benDe(a.id).filter((x) => ['saude', 'odonto'].includes(benCat(x.beneficio).familia) && !['recusado', 'cancelado', 'previsto'].includes(x.status)).forEach((x) => {
      const faltaT = [!a.matricula && 'matrícula', !a.cpf && 'CPF', !a.nasc && 'nascimento', !a.sexo && 'sexo', !a.estado_civil && 'estado civil'].filter(Boolean);
      linhas.push([a.matricula || '', titulo(a.nome), a.cpf || '', a.nasc || '', a.sexo || '', a.estado_civil || '', 'Titular', benCat(x.beneficio).nome, x.plano || '', a.local, a.cc_codigo || '', dm(a.data_inicio), (BEN_ST[x.status] || [])[2], x.wkf_ref || '', [faltaT.length ? 'falta: ' + faltaT.join(', ') : '', a.era_dependente_cimed ? 'era dependente: mover para titular' : ''].filter(Boolean).join(' · ')]);
      depsDe(a.id).filter((d) => (benCat(x.beneficio).familia === 'saude' ? d.plano_saude : d.odonto)).forEach((d) => {
        const faltaD = [!d.cpf && 'CPF', !d.nasc && 'nascimento'].filter(Boolean);
        const regra = (d.parentesco === 'Filho(a)' || d.parentesco === 'Enteado(a)') && d.idade >= 21 ? (d.idade >= 24 ? 'acima de 24: não entra' : d.universitario ? 'universitário: declaração semestral' : 'de 21 a 24: só universitário') : ['Pai', 'Mãe'].includes(d.parentesco) && benCat(x.beneficio).codigo !== 'odonto_sulamerica' ? 'pais: só no odonto SulAmérica' : '';
        linhas.push([a.matricula || '', titulo(d.nome), d.cpf || '', d.nasc || '', '', '', d.parentesco, benCat(x.beneficio).nome, x.plano || '', a.local, a.cc_codigo || '', dm(a.data_inicio), 'dependente', x.wkf_ref || '', [faltaD.length ? 'falta: ' + faltaD.join(', ') : '', regra].filter(Boolean).join(' · ')]);
      });
    }));
    return baixarCsvSimples(`conferencia-wkf-saude-odonto-${sufixo}.csv`, ['Matrícula', 'Nome', 'CPF', 'Nascimento', 'Sexo', 'Estado civil', 'Parentesco', 'Benefício', 'Plano', 'Local', 'Centro de custo', 'Integração', 'Situação', 'WKF', 'Alertas'], linhas);
  }
  if (k === 'fretado') return baixarCsvSimples(`fretado-${sufixo}.csv`, ['Matrícula', 'Nome', 'Unidade', 'Turno', 'Código da linha no LG', 'Linha', 'Itinerário (turno)', 'Chegada', 'Telefone', 'Integração', 'Situação'],
    ps.filter((a) => benDe(a.id).some((x) => x.beneficio === 'fretado' && !['recusado', 'cancelado'].includes(x.status))).map((a) => { const it = itinerariosDe(a)[0]; return [a.matricula || '', titulo(a.nome), a.local, a.turno || '', a.fretado_codigo || '', a.fretado || '', it ? it.turno : '', it ? it.chegada || '' : '', a.telefone || '', dm(a.data_inicio), (BEN_ST[benDe(a.id).find((x) => x.beneficio === 'fretado').status] || [])[2]]; }));
  if (k === 'vt') return baixarCsvSimples(`vale-transporte-${sufixo}.csv`, ['Matrícula', 'Nome', 'CPF', 'Nascimento', 'Local', 'Integração', 'Meio (kit)', 'Passagens por dia (kit)', 'Valor da passagem (kit)', 'Situação'],
    ps.filter((a) => benDe(a.id).some((x) => x.beneficio === 'vt' && !['recusado', 'cancelado'].includes(x.status))).map((a) => { const o = (opcDe(a.id) || {}).dados || {}; return [a.matricula || '', titulo(a.nome), a.cpf || '', a.nasc || '', a.local, dm(a.data_inicio), o.vt_meio || '', o.vt_qtd || '', o.vt_valor || '', (BEN_ST[benDe(a.id).find((x) => x.beneficio === 'vt').status] || [])[2]]; }));
  return baixarCsvSimples(`beneficios-checklist-${sufixo}.csv`, ['Matrícula', 'Nome', 'Local', 'Grupo', 'Integração', 'Benefício', 'Situação', 'Plano', 'Pendência', 'Atualizado por', 'Em'],
    ps.flatMap((a) => benDe(a.id).map((x) => [a.matricula || '', titulo(a.nome), a.local, grupoBen(a), dm(a.data_inicio), benCat(x.beneficio).nome, (BEN_ST[x.status] || [])[2], x.plano || '', x.pendencia || '', x.atualizado_por || '', quando(x.atualizado_em)])));
}

// --------------------------------------------------------- especialistas
VIEWS.especialistas = () => {
  const g = Object.keys(GATILHO);
  const unidades = ['todas', ...E.locais.map((l) => l.nome), 'fabrica', 'cd', 'sp', 'comercial', 'pa_sp_comercial', 'ssbv_cds'];
  const optU = (sel) => unidades.map((u) => `<option value="${esc(u)}" ${u === sel ? 'selected' : ''}>${esc(GRUPO_TURMA[u] || ({ todas: 'Todas as unidades', fabrica: 'Fábricas', cd: 'CDs', sp: 'Escritórios SP', comercial: 'Comercial' })[u] || u)}</option>`).join('');
  const linhaBp = (b, i) => `<div class="area" data-bp="${i}"><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(b.unidade)}</select></div><div class="ac"><label class="olho">Setor (opcional)</label><input class="sel-doc" data-f="area" value="${esc(b.area || '')}" placeholder="vazio = unidade inteira"></div><div class="ac"><label class="olho">BP</label><input class="sel-doc" data-f="nome" value="${esc(b.nome || '')}" placeholder="Nome"><input class="sel-doc" data-f="email" value="${esc(b.email || '')}" placeholder="nome@grupocimed.com.br" style="margin-top:4px"></div><div class="linha-flex"><button class="bt peq" data-act="bpSalvar" data-i="${i}">Salvar</button>${b.id ? `<button class="bt peq fan" data-act="bpApagar" data-id="${b.id}">Apagar</button>` : ''}</div></div>`;
  const linhaAreaDest = (d, i) => `<div class="area" data-adest="${i}"><div class="ac"><label class="olho">Área</label><select class="sel-doc" data-f="area_id">${E.areas.map((a) => `<option value="${esc(a.id)}" ${a.id === d.area_id ? 'selected' : ''}>${esc(a.nome)}</option>`).join('')}</select></div><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(d.unidade)}</select></div><div class="ac"><label class="olho">E-mails</label><input class="sel-doc" style="width:100%;min-height:40px" data-f="emails" value="${esc((d.emails || []).join(', '))}" placeholder="nome@grupocimed.com.br"></div><div class="linha-flex"><button class="bt peq" data-act="adestSalvar" data-i="${i}">Salvar</button>${d.id ? `<button class="bt peq fan" data-act="adestApagar" data-id="${d.id}">Apagar</button>` : ''}</div></div>`;
  const linhaDest = (d, i) => `<div class="area" data-dest="${i}"><div class="ac"><label class="olho">Unidade</label><select class="sel-doc" data-f="unidade">${optU(d.unidade)}</select><input class="sel-doc" data-f="descricao" value="${esc(d.descricao || '')}" placeholder="Descrição" style="margin-top:4px"></div><div class="ac" style="grid-column:span 2"><label class="olho">E-mails e listas</label><input class="sel-doc" style="width:100%;min-height:40px" data-f="emails" value="${esc((d.emails || []).join(', '))}" placeholder="lista@grupocimed.com.br, portaria@grupocimed.com.br"><label class="marcas" style="margin-top:4px"><label><input type="checkbox" data-f="copiar_requisitantes" ${d.copiar_requisitantes !== false ? 'checked' : ''}> Requisitantes da turma em cópia</label></label></div><div class="linha-flex"><button class="bt peq" data-act="destSalvar" data-i="${i}">Salvar</button>${d.id ? `<button class="bt peq fan" data-act="destApagar" data-id="${d.id}">Apagar</button>` : ''}</div></div>`;
  return `<div class="cab"><div class="t"><span class="olho">Comunicação</span><h1>Especialistas e avisos</h1><p class="apoio">Quem recebe e-mail, em que momento e com quais dados. Área sem e-mail cadastrado não recebe nada. As marcadas como sugestão vieram do mapeamento e precisam da sua validação antes de ligar. Separe vários e-mails por vírgula. Gestor, analista de R&S e BP têm destino por candidato: o e-mail do requisitante na ficha, o e-mail da analista no cadastro de pessoas e a tabela de BPs abaixo; a lista da área é só reserva.</p></div></div>
  ${g.map((k) => `<div class="caixa" style="padding:0"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>${esc(GATILHO[k])}</h2></div>
   ${E.areas.filter((a) => a.gatilho === k).map((a) => `<div class="area"><div><b>${esc(a.nome)}</b><div class="linha-flex" style="gap:6px;margin-top:3px">${a.origem === 'sugestao' ? chip('c-ama', '', 'Sugestão') : chip('c-ag', 'ok', 'Você pediu')}<span class="apoio" style="font-size:12px">${esc(CONDICAO[a.condicao] || a.condicao)}</span></div><div class="apoio" style="font-size:12.5px;margin-top:4px">${esc(a.leva)}</div></div>
     <div class="ac" style="grid-column:span 2"><label class="olho" for="em_${esc(a.id)}">E-mails${['gestor', 'rs', 'bp'].includes(a.id) ? ' (reserva)' : ''}</label><input id="em_${esc(a.id)}" class="sel-doc" style="width:100%;min-height:40px" value="${esc((a.emails || []).join(', '))}" placeholder="nome@grupocimed.com.br"></div>
     <div class="linha-flex"><button class="troca" data-act="togArea" data-id="${esc(a.id)}" aria-pressed="${a.ativo}" aria-label="Ligar ${esc(a.nome)}"></button><button class="bt peq" data-act="salvarArea" data-id="${esc(a.id)}">Salvar</button></div></div>`).join('')}</div>`).join('')}
  ${E.areaDest ? `<div class="caixa" style="padding:0" id="areaDestBox"><div style="padding:14px 16px;border-bottom:1px solid var(--linha)"><h2>Destino por unidade</h2><p class="apoio" style="font-size:12.5px">Para a área que tem uma pessoa em cada unidade (por exemplo, Facilities em Pouso Alegre, Fernão Dias e SSBV). Com linha para a unidade, o aviso vai para estes e-mails; sem linha, para a lista da área acima. Frotas e Facilities também recebem um novo e-mail quando cargo, setor, centro de custo, gestor, unidade, turma, turno, camisa ou uso do carro mudam depois do aviso, e quando a admissão é cancelada.</p></div>
   ${E.areaDest.map(linhaAreaDest).join('')}${linhaAreaDest({ area_id: 'facilities', unidade: 'Pouso Alegre' }, E.areaDest.length)}</div>` : ''}
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
    case 'turma': if (await tentar(() => api.atualizar(id, { data_inicio: el.dataset.t, turma_id: el.dataset.tid || null }), 'Turma alterada')) { E.turma = el.dataset.t; await recarregar(); abrirFicha(id); } break;
    case 'salvarAso': {
      const st = $('#asoSel').value;
      const p = { aso_status: st, aso_data: $('#asoData').value || null, aso_hora: ($('#asoHora')?.value || '').trim() || null, aso_local: ($('#asoLocal')?.value || '').trim() || null, aso_restricao: st === 'apto_restricao' ? ($('#asoRestr')?.value || '').trim() || null : null };
      if ($('#asoObs')) { p.aso_obs = valLista('asoObs') || null; p.aso_responsavel = valLista('asoResp') || null; }
      if (st === 'agendado' && !p.aso_data) { toast('Informe a data do exame.', true); return; }
      if (st === 'apto_restricao' && !p.aso_restricao) { toast('Descreva a restrição.', true); return; }
      if ((st === 'apto' || st === 'apto_restricao') && !asoOk(adm(id) || {}) && E.params?.vacina_bloqueia_apto !== 'nao' && adm(id)?.vacina_status && adm(id).vacina_status !== 'comprovada') { toast('O ASO apto fica bloqueado até a vacinação ser comprovada. Registre as vacinas antes.', true); return; }
      if (await tentar(() => api.atualizar(id, p), st === 'agendado' && adm(id)?.email ? 'ASO agendado. O candidato recebe o e-mail com data, hora e local.' : 'ASO atualizado')) { await recarregar(); if ($('.gaveta')) abrirFicha(id); else render(); }
      break;
    }
    case 'osAbrir': abrirOs(id); break;
    // agenda da Helaine (SQL 27)
    case 'iniLimpar': E.iniF = { setor: '', local: '', mes: '' }; render(); break;
    case 'confSalvar': if (await tentar(() => api.atualizar(id, { conferente_docs: valLista('confSel') || null }), 'Responsável pela conferência salva')) await aposSaude(id); break;
    case 'docsRevisar': if (await tentar(() => api.docsRevisar(id), 'Revisão final registrada: documentos ok')) await aposSaude(id); break;
    case 'docsRevisarDesfazer': if (confirm('Desfazer a revisão final dos documentos?') && await tentar(() => api.docsRevisar(id, true), 'Revisão desfeita')) await aposSaude(id); break;
    case 'sindConferir': { const s = ($('#sindIn')?.value || '').trim(); if (!s) { toast('Escolha o sindicato.', true); return; } if (await tentar(() => api.sindicatoConferir(id, s, ($('#sindObs')?.value || '').trim()), 'Sindicato conferido')) await aposSaude(id); break; }
    case 'sindConferirLinha': { const s = (document.querySelector(`[data-sind="${id}"]`)?.value || '').trim(); if (!s) { toast('Escolha o sindicato.', true); return; } if (await tentar(() => api.sindicatoConferir(id, s, null), 'Sindicato conferido')) await recarregar(); break; }
    case 'cobrarLg': if (await tentar(() => api.cobrarLg(id), 'Cobrança enviada à recrutadora (e-mail e sino)')) await aposSaude(id); break;
    case 'turmaConcluida': {
      const tu = E.turmasCal.find((x) => x.id === id);
      if (!confirm('Enviar a todas as áreas o e-mail "Admissões concluídas" da turma de ' + dm(tu?.data) + '? Vai a lista de quem tem matrícula.')) return;
      const r = await tentar(() => api.avisarTurmaConcluida(id, true));
      if (r === false) toast('Ninguém com matrícula nesta turma ainda.', true); else if (r) { toast('E-mail às áreas na fila'); await recarregar(); }
      break;
    }
    case 'excCancelar': if (confirm('Cancelar esta liberação?') && await tentar(() => api.excecaoCancelar(+id), 'Liberação cancelada')) await recarregar(); break;
    case 'listaTipo': E.listaTipo = el.dataset.t; render(); break;
    case 'listaSalvar': {
      const box = el.closest('.area'); const val = box.querySelector('[data-f="valor"]').value.trim();
      if (!val) { toast('Escreva o item.', true); return; }
      const o = { tipo: E.listaTipo, valor: val, filtro: box.querySelector('[data-f="filtro"]')?.value || null, ativo: box.querySelector('[data-f="ativo"]').checked };
      if (box.dataset.lid) o.id = +box.dataset.lid;
      if (await tentar(() => api.salvarLista(o), 'Lista salva')) await recarregar();
      break;
    }
    case 'listaApagar': if (confirm('Apagar este item? Prefira desmarcar "Ativo" se ele já foi usado.') && await tentar(() => api.apagarLista(+id), 'Item apagado')) await recarregar(); break;
    case 'regraSalvar': {
      const box = el.closest('.area'); const g = (f) => (box.querySelector(`[data-f="${f}"]`)?.value || '').trim();
      if (!g('sindicato')) { toast('Informe o sindicato.', true); return; }
      const o = { local: g('local'), cargo_contem: g('cargo_contem') || null, sindicato: g('sindicato'), observacao: g('observacao') || null, ativo: true };
      if (box.dataset.rid) o.id = +box.dataset.rid;
      if (await tentar(() => api.salvarRegraSindicato(o), 'Regra salva')) await recarregar();
      break;
    }
    case 'regraApagar': if (confirm('Apagar esta regra de sindicato?') && await tentar(() => api.apagarRegraSindicato(+id), 'Regra apagada')) await recarregar(); break;
    // Benefícios (SQL 28)
    case 'benAbrir': abrirBen(id); break;
    case 'benSalvar': {
      const b = el.dataset.b; const box = el.closest('[data-benbox]'); const st = box.querySelector('[data-f="st"]').value;
      if (!st || st === 'previsto') { toast('Escolha a situação. Antes da matrícula o item fica como previsão.', true); return; }
      const pend = box.querySelector('[data-f="pend"]').value.trim(); const plano = box.querySelector('[data-f="plano"]').value.trim();
      const antes = benDe(id).find((x) => x.beneficio === b) || {};
      let wkf = null;
      if (st === 'enviado' && antes.status !== 'enviado' && temSql31() && ['saude', 'odonto'].includes(benCat(b).familia)) { wkf = await perguntar('Em qual WKF foi enviado?', 'Ex.: WKF Unimed 14/10/2026 (deixe vazio se não souber)', ''); if (wkf === null) return; }
      if (await tentar(() => api.beneficioAtualizar([id], b, st, pend, plano && plano !== (antes.plano || '') ? plano : null, wkf), 'Benefício atualizado')) { await recarregar(); abrirBen(id); }
      break;
    }
    case 'benLote': {
      const ids = [...document.querySelectorAll('[data-bsel]:checked')].map((c) => c.dataset.bsel);
      if (!ids.length) { toast('Marque as pessoas na tabela.', true); return; }
      const b = $('#bLoteBen').value; const st = $('#bLoteSt').value;
      let wkf = null;
      if (st === 'enviado' && temSql31() && ['saude', 'odonto'].includes(benCat(b).familia)) { wkf = await perguntar('Em qual WKF foram enviados?', 'Ex.: WKF Unimed 14/10/2026 (deixe vazio se não souber)', ''); if (wkf === null) return; }
      const n = await tentar(() => api.beneficioAtualizar(ids, b, st, null, null, wkf));
      if (n != null) { toast(plural(n, 'item atualizado', 'itens atualizados') + (n < ids.length ? ' (quem não tem este benefício ficou de fora)' : '')); await recarregar(); }
      break;
    }
    case 'benCsv': if (el.dataset.k === 'fretado') await garantirItins(); exportarBen(el.dataset.k); break;
    case 'benItinFiltro': E.itinUni = $('#itUni').value; E.itinTurno = $('#itTurno').value; render(); break;
    // fluxo da Saúde (SQL 26)
    case 'socMarcar': if (await tentar(async () => { if (!(await api.socMarcar(id))) throw new Error('Não foi possível marcar: a admissão foi cancelada.'); return true; }, adm(id)?.email ? 'Cadastrado no SOC. O candidato recebe o e-mail da etapa de exame médico.' : 'Cadastrado no SOC (candidato sem e-mail na ficha)')) await aposSaude(id); break;
    case 'socDesmarcar': if (confirm('Desmarcar o cadastro no SOC?') && await tentar(() => api.socMarcar(id, true), 'Cadastro no SOC desmarcado')) await aposSaude(id); break;
    case 'saudeContato': if (await tentar(() => api.saudeAcao(id, 'contato'), 'Contato registrado')) await aposSaude(id); break;
    case 'saudeSemContato': if (await tentar(() => api.saudeAcao(id, 'sem_contato'), 'Registrado: tentativas por até ' + (E.params?.saude_janela_contato_horas || '24') + 'h')) await aposSaude(id); break;
    case 'saudeFalta': if (confirm('Registrar que ' + titulo(adm(id)?.nome || '') + ' faltou ao exame?') && await tentar(() => api.saudeAcao(id, 'falta'), 'Falta registrada: tentativas por até ' + (E.params?.saude_janela_contato_horas || '24') + 'h')) await aposSaude(id); break;
    case 'saudeVacina': {
      const v = $('#vacSel').value; const obs = ($('#vacObs')?.value || '').trim();
      if (v === 'recusada' && !confirm('Recusa de vacinação bloqueia o ASO e vai para o R&S decidir. Confirmar?')) return;
      if (await tentar(() => api.saudeAcao(id, 'vacina', v, obs), v === 'recusada' ? 'Recusa registrada e enviada ao R&S' : 'Vacinação registrada')) await aposSaude(id);
      break;
    }
    case 'saudeCompl': if (await tentar(() => api.saudeAcao(id, 'complementares', $('#complSel').value || null), 'Prazo do ASO atualizado')) await aposSaude(id); break;
    case 'saudeModo': if (await tentar(() => api.saudeAcao(id, 'modo', $('#modoSel').value || null, ($('#modoClinica')?.value || '').trim()), 'Local do exame registrado')) await aposSaude(id); break;
    case 'saudeEscalar': {
      const m = $('#escMotivo').value;
      if (!confirm('Enviar ao R&S (' + MOTIVO_SAUDE[m] + ')? A analista recebe o e-mail e decide na ficha.')) return;
      if (await tentar(() => api.saudeAcao(id, 'escalar', m, ($('#escTexto')?.value || '').trim()), 'Enviado ao R&S')) await aposSaude(id);
      break;
    }
    case 'saudeErroSoc': {
      const tx = ($('#socErro')?.value || '').trim();
      if (!tx) { toast('Descreva o que está errado no SOC.', true); return; }
      if (await tentar(() => api.saudeAcao(id, 'erro_soc', null, tx), 'Erro enviado ao R&S')) await aposSaude(id);
      break;
    }
    case 'rsSaude': {
      const d = el.dataset.d; const nome = titulo(adm(id)?.nome || '');
      if (d === 'declinar' && !confirm('Declinar a vaga de ' + nome + '? As áreas recebem o aviso de cancelamento.')) return;
      if (await tentar(() => api.rsDecidirSaude(id, d), d === 'declinar' ? 'Vaga declinada' : d === 'soc_corrigido' ? 'SOC corrigido: a Saúde segue' : 'Nova tentativa: a Saúde retoma o exame')) await aposSaude(id);
      break;
    }
    case 'saudeParam': {
      const lista = ($('#vacLista')?.value || '').trim(); const bloq = $('#vacBloq').value;
      if (!lista) { toast('Informe ao menos uma vacina.', true); return; }
      if (await tentar(async () => { await api.saudeParametro('vacinas_obrigatorias', lista); return api.saudeParametro('vacina_bloqueia_apto', bloq); }, 'Vacinas obrigatórias salvas')) { await recarregar(); render(); }
      break;
    }
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
    case 'adestSalvar': {
      const box = document.querySelector(`[data-adest="${el.dataset.i}"]`); const atual = E.areaDest[+el.dataset.i];
      const o = atual ? { id: atual.id } : {};
      box.querySelectorAll('[data-f]').forEach((x) => { o[x.dataset.f] = x.value.trim(); });
      o.emails = o.emails.split(/[,;\s]+/).map((x) => x.trim().toLowerCase()).filter(Boolean);
      const ruins = o.emails.filter((x) => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
      if (!o.emails.length || ruins.length) { toast(ruins.length ? 'E-mail inválido: ' + ruins.join(', ') : 'Informe ao menos um e-mail.', true); return; }
      if (await tentar(() => api.salvarAreaDestino(Object.assign({ ativo: true }, o)), 'Destino salvo')) await recarregar(); break;
    }
    case 'adestApagar': if (await tentar(() => api.apagarAreaDestino(+id || id), 'Destino removido')) await recarregar(); break;
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

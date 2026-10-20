// =====================================================================
// CIMED · Plataforma de Admissão
// ÚNICO ARQUIVO QUE VOCÊ PRECISA EDITAR
// =====================================================================

// Supabase da ADMISSÃO (projeto novo, não o do Gestão de Projetos).
// Project Settings > Data API: copie a URL e a chave publishable (anon).
// A chave pode ficar visível: quem protege os dados é a RLS do 01.
// NUNCA coloque aqui a service_role.
export const SUPABASE_URL = 'https://wtkqcrccrxhobhjifqdj.supabase.co';
export const SUPABASE_ANON_KEY = 'sb_publishable_-qdtlNmea5BWtFhvDibpLA_zsf29DzO';

// Login com a conta Microsoft da Cimed. Deixe false até a TI criar o
// app no Entra e você ligar o provedor Azure no Supabase (ver README).
export const LOGIN_MICROSOFT = false;

// Endereço publicado da Esteira: é o que vai no link e no QR do kit on-line.
// A prévia do Lovable (preview--...) é bloqueada pelo próprio Lovable, então
// o link nunca pode apontar para ela.
export const SITE_PUBLICO = 'https://admissao.rhcimed.online/app/esteira.html';

// Quem acessa e com qual papel NÃO fica aqui: fica na tabela "perfis"
// do Supabase. Conta fora da tabela não vê nada.

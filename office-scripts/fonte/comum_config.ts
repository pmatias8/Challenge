// ---------- CONFIGURE AQUI ----------
const API_BASE = "https://vmpay-proxy.SEU-USUARIO.workers.dev/api/v1";
// O script fica guardado no SEU OneDrive. Se você adicionar um botão deste
// script numa planilha compartilhada, o script (e este token) passa a ser
// compartilhado com quem abre a planilha. Nesse caso deixe vazio e digite o
// token no campo "token" ao rodar.
const TOKEN_VMPAY = "";
// ------------------------------------

const OFFSET_UTC = -3;
const PER_PAGE = 1000;
const LOTE = 8;                 // páginas pedidas em paralelo


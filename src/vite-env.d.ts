/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_AAD_CLIENT_ID: string;
  readonly VITE_AAD_TENANT: string;
  readonly VITE_GRAPH_BASE_URL?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}

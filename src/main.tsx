import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { pca } from './auth/msalConfig';
import { AuthProvider } from './auth/AuthProvider';
import { DataSourceProvider } from './app/DataSourceProvider';
import { installDebugAudit } from './app/debugAudit';
import { installDebugLog } from './app/debugLog';
import { installTracer } from './app/tracer';
import { App } from './App';

const queryClient = new QueryClient({
  defaultOptions: {
    queries: { retry: 1, refetchOnWindowFocus: false },
  },
});

// msal-browser v3 requires explicit initialization before any other API call.
async function bootstrap() {
  // Wire the Graph call-log inspector (window.__gsaAudit) before any call runs.
  installDebugAudit();
  // Wire policy-evaluation debug logging (window.__gsaDebug): CA matching + What-If.
  installDebugLog();
  // Wire the general structured trace facility (window.__gsaTrace).
  installTracer();

  await pca.initialize();

  const handle = await pca.handleRedirectPromise();
  if (handle?.account) pca.setActiveAccount(handle.account);

  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <FluentProvider theme={webLightTheme} style={{ height: '100vh' }}>
        <QueryClientProvider client={queryClient}>
          <AuthProvider>
            <DataSourceProvider>
              <App />
            </DataSourceProvider>
          </AuthProvider>
        </QueryClientProvider>
      </FluentProvider>
    </StrictMode>,
  );
}

void bootstrap();

import { spawn } from 'node:child_process';

const port = 9876 + Math.floor(Math.random() * 1000);
const baseUrl = `http://127.0.0.1:${port}`;

async function main() {
  const server = spawn(process.execPath, ['--import', 'tsx', 'server/index.ts'], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      PORT: String(port),
      HOST: '127.0.0.1',
      NODE_ENV: 'production',
      ALLOW_UNAUTHENTICATED_CAPTION_TESTING: 'false',
      SUPABASE_SERVICE_ROLE_KEY: 'your_service_role_key',
      GEMINI_API_KEY: 'your_backend_only_gemini_key',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  const output: string[] = [];
  server.stdout.on('data', (chunk) => output.push(String(chunk)));
  server.stderr.on('data', (chunk) => output.push(String(chunk)));

  try {
    await waitForHealth();

    const health = await fetchJson('/api/health');
    assert(health.status === 200, `Expected health status 200, got ${health.status}`);
    assert(JSON.stringify(health.body) === '{"ok":true}', `Health leaked internal metadata: ${JSON.stringify(health.body)}`);

    const admin = await fetchJson('/api/admin/usage');
    assert(admin.status !== 200, `Admin usage endpoint was directly accessible: ${JSON.stringify(admin.body)}`);

    const captions = await fetchJson('/api/generate-captions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ elements: [] }),
    });
    assert(captions.status !== 200, `Unauthenticated captions endpoint was directly accessible: ${JSON.stringify(captions.body)}`);

    console.log('security-smoke-test passed');
  } finally {
    server.kill();
  }

  async function waitForHealth(): Promise<void> {
    const startedAt = Date.now();
    while (Date.now() - startedAt < 10_000) {
      if (server.exitCode !== null) throw new Error(`Server exited early:\n${output.join('')}`);
      try {
        const response = await fetch(`${baseUrl}/api/health`);
        if (response.ok) return;
      } catch {
        await new Promise((resolve) => setTimeout(resolve, 150));
      }
    }
    throw new Error(`Timed out waiting for server:\n${output.join('')}`);
  }
}

async function fetchJson(path: string, init?: RequestInit): Promise<{ status: number; body: unknown }> {
  const response = await fetch(`${baseUrl}${path}`, init);
  const contentType = response.headers.get('content-type') ?? '';
  return {
    status: response.status,
    body: contentType.includes('application/json') ? await response.json() : await response.text(),
  };
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

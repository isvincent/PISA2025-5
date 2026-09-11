import fs from 'node:fs';
import path from 'node:path';

for (const entrypoint of ['index.html', 'dist/index.html']) {
  const html = fs.readFileSync(entrypoint, 'utf8');
  const script = html.match(/<script>([\s\S]*?)<\/script>/);
  if (!script) throw new Error(`${entrypoint}: missing inline script`);
  new Function(script[1]);

  const base = path.dirname(entrypoint);
  for (const match of html.matchAll(/(?:src|href)="([^"]+)"/g)) {
    const value = match[1];
    if (!/^(https?:|#|data:)/.test(value) && !fs.existsSync(path.join(base, value))) {
      throw new Error(`${entrypoint}: missing local asset: ${value}`);
    }
  }
}

console.log('GitHub Pages and Sites entrypoints are valid');

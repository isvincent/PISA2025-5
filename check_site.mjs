import fs from 'node:fs';
const html = fs.readFileSync('dist/index.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/);
if (!script) throw new Error('Missing inline script');
new Function(script[1]);
for (const match of html.matchAll(/(?:src|href)="([^"]+)"/g)) {
  const value = match[1];
  if (!/^(https?:|#|data:)/.test(value) && !fs.existsSync(`dist/${value}`)) {
    throw new Error(`Missing local asset: ${value}`);
  }
}
console.log('HTML assets and JavaScript OK');

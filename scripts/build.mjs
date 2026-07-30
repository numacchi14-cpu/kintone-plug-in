import { rm } from 'node:fs/promises';
import { resolve } from 'node:path';
import { build } from 'vite';

const root = process.cwd();

await rm(resolve(root, 'dist'), { recursive: true, force: true });

async function buildEntry(entry, name, fileName) {
  await build({
    configFile: false,
    root,
    build: {
      outDir: 'dist',
      emptyOutDir: false,
      target: 'es2020',
      sourcemap: true,
      lib: {
        entry: resolve(root, entry),
        formats: ['iife'],
        name,
        fileName: () => fileName
      },
      rollupOptions: {
        output: {
          inlineDynamicImports: true
        }
      }
    }
  });
}

await buildEntry('src/desktop.ts', 'KintoneExcelReportDesktop', 'js/desktop.js');
await buildEntry('src/config.ts', 'KintoneExcelReportConfig', 'js/config.js');

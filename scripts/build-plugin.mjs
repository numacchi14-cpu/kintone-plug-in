import { copyFile, cp, mkdir, readFile, readdir, rm } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';

const root = process.cwd();
const packageDir = resolve(root, 'plugin-package');
const packageJson = JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8'));
const version = packageJson.version;
const versionedZipPath = resolve(root, 'dist', `kintone-excel-report-plugin-v${version}.zip`);
const latestZipPath = resolve(root, 'dist', 'kintone-excel-report-plugin.zip');
const privateKeyDir = resolve(root, 'private');
const privateKeyPath = resolve(privateKeyDir, 'plugin.ppk');

await rm(packageDir, { recursive: true, force: true });
await mkdir(resolve(packageDir, 'js'), { recursive: true });
await mkdir(resolve(packageDir, 'html'), { recursive: true });
await mkdir(resolve(packageDir, 'css'), { recursive: true });

await cp(resolve(root, 'plugin', 'manifest.json'), resolve(packageDir, 'manifest.json'));
await cp(resolve(root, 'plugin', 'image'), resolve(packageDir, 'image'), { recursive: true });
await cp(resolve(root, 'dist', 'js'), resolve(packageDir, 'js'), { recursive: true });
await cp(resolve(root, 'public', 'config.html'), resolve(packageDir, 'html', 'config.html'));
await cp(resolve(root, 'public', 'css'), resolve(packageDir, 'css'), { recursive: true });

const packerCli = resolve(root, 'node_modules', '@kintone', 'plugin-packer', 'bin', 'cli.js');
if (existsSync(packerCli)) {
  await mkdir(privateKeyDir, { recursive: true });

  const packerArgs = existsSync(privateKeyPath)
    ? [packerCli, packageDir, '--ppk', privateKeyPath, '--out', versionedZipPath]
    : [packerCli, packageDir, '--out', versionedZipPath];

  const result = spawnSync(process.execPath, packerArgs, {
    cwd: root,
    stdio: 'inherit',
    shell: false
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }

  if (!existsSync(privateKeyPath)) {
    const distEntries = await readdir(resolve(root, 'dist'));
    const generatedKey = distEntries.find((entry) => entry.endsWith('.ppk'));

    if (generatedKey) {
      await copyFile(resolve(root, 'dist', generatedKey), privateKeyPath);
      console.log(`Saved plugin private key: ${privateKeyPath}`);
    } else {
      console.warn('Private key was not found. Future packages may not update the same kintone plugin.');
    }
  }

  await copyFile(versionedZipPath, latestZipPath);
  console.log(`Versioned plugin package: ${versionedZipPath}`);
  console.log(`Latest plugin package: ${latestZipPath}`);
} else {
  console.warn('kintone-plugin-packer was not found. Staged plugin files were written to plugin-package/.');
}

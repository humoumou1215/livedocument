import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs/promises';

export const DEFAULT_PORT = 18765;
export function stateDirectory(env = process.env) {
  if (env.LIVEDOC_HOME) return path.resolve(env.LIVEDOC_HOME);
  return path.join(env.XDG_RUNTIME_DIR || path.join(os.homedir(), '.livedocument'), env.XDG_RUNTIME_DIR ? 'livedocument' : '');
}
export function statePath(env = process.env) { return path.join(stateDirectory(env), 'runtime.json'); }
export async function saveState(value, env = process.env) {
  const dir = stateDirectory(env);
  await fs.mkdir(dir, {recursive:true, mode:0o700});
  if (process.platform !== 'win32') await fs.chmod(dir, 0o700);
  const file = statePath(env);
  const temp = `${file}.${process.pid}.tmp`;
  await fs.writeFile(temp, JSON.stringify(value), {encoding:'utf8', mode:0o600, flag:'wx'});
  if (process.platform !== 'win32') await fs.chmod(temp, 0o600);
  await fs.rename(temp, file);
}
export async function loadState(env = process.env) {
  const data = await fs.readFile(statePath(env), 'utf8');
  const obj = JSON.parse(data);
  if (!obj || !Number.isInteger(obj.port) || !/^[a-f0-9]{64}$/.test(obj.token) || typeof obj.pid !== 'number') throw new Error('Invalid runtime state');
  return obj;
}
export async function deleteStateIfOwned(token, env = process.env) {
  try { if ((await loadState(env)).token === token) await fs.unlink(statePath(env)); } catch (e) { if (e.code !== 'ENOENT') throw e; }
}

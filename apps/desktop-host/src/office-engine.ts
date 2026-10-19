/** Resolve packaged Office engine manifests from their complete, unpacked resource directories. */
import { createRequire, registerHooks, type ModuleHooks } from 'node:module'
import { realpathSync } from 'node:fs'
import type fs from 'node:fs'
import { basename, dirname, join, relative } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

/**
 * Locate the archive containing a packaged runtime.
 * @param runtimeDir - Prepared or ASAR-contained runtime directory.
 * @returns Parent archive path, or undefined for a prepared directory.
 */
export function runtimeArchivePath(runtimeDir: string): string | undefined {
  const parent = dirname(runtimeDir)
  return basename(parent) === 'app.asar' ? parent : undefined
}

/**
 * Keep engine executable and resource paths usable by native child processes outside Electron.
 * Hooks apply only to this thread; worker threads must install their own resolver.
 * @param runtimeDir - Prepared or ASAR-contained dsh runtime directory.
 * @returns Installed resolver for the Host lifetime, or undefined for a non-ASAR runtime.
 */
export function installOfficeEngineResolution(runtimeDir: string): ModuleHooks | undefined {
  if (runtimeArchivePath(runtimeDir) === undefined) return undefined
  const root = realpathSync(runtimeDir)
  const archive = dirname(root)
  const source = pathToFileURL(join(root, 'node_modules', '@deepseek-ai', 'libreoffice-kit-')).href
  const destination = pathToFileURL(join(`${archive}.unpacked`, relative(archive, root), 'node_modules', '@deepseek-ai', 'libreoffice-kit-')).href

  const require = createRequire(import.meta.url)
  const cjsFs = require('node:fs') as typeof fs
  const origLstat = cjsFs.lstatSync
  cjsFs.lstatSync = ((path: fs.PathLike, options?: Parameters<typeof origLstat>[1]) => {
    const result = origLstat(path, options)
    return result === null ? undefined : result
  }) as typeof origLstat
  const origStat = cjsFs.statSync
  cjsFs.statSync = ((path: fs.PathLike, options?: Parameters<typeof origStat>[1]) => {
    const result = origStat(path, options)
    return result === null ? undefined : result
  }) as typeof origStat
  const plainFs = require('fs') as typeof fs
  plainFs.lstatSync = cjsFs.lstatSync
  plainFs.statSync = cjsFs.statSync

  const fsWrapperUrl = 'data:text/javascript,' + encodeURIComponent(
    'import * as fs from "node:fs";' +
    'export * from "node:fs";' +
    'export const lstatSync = (...args) => { const r = fs.lstatSync(...args); return r === null ? undefined : r; };' +
    'export const statSync = (...args) => { const r = fs.statSync(...args); return r === null ? undefined : r; };' +
    'export default (fs.default || fs);',
  )

  const hook = registerHooks({
    resolve(specifier, context, nextResolve) {
      if ((specifier === 'node:fs' || specifier === 'fs') && /(?:@deepseek-ai|%40deepseek-ai)\/libreoffice-kit/u.test(context.parentURL ?? '')) {
        return { url: fsWrapperUrl, shortCircuit: true }
      }
      const resolved = nextResolve(specifier, context)
      if (!resolved.url.startsWith('file:')) return resolved
      const engineRequest = /^@deepseek-ai\/libreoffice-kit-(?:darwin|win32|linux)-/u.test(specifier)
      const engineTarget = /\/node_modules\/@deepseek-ai\/libreoffice-kit-(?:darwin|win32|linux)-[^/]+\//u
        .test(new URL(resolved.url).pathname)
      if (!engineRequest && !engineTarget) return resolved
      const canonical = pathToFileURL(realpathSync(fileURLToPath(resolved.url))).href
      if (!canonical.startsWith(source)) {
        if (canonical.startsWith(pathToFileURL(archive + '/').href)) {
          throw new Error(`desktop Office engine resolved outside the runtime package directory: ${resolved.url}`)
        }
        return resolved
      }
      const physical = realpathSync(fileURLToPath(destination + canonical.slice(source.length)))
      return { ...resolved, url: pathToFileURL(physical).href }
    },
  })

  return {
    deregister() {
      cjsFs.lstatSync = origLstat
      cjsFs.statSync = origStat
      plainFs.lstatSync = origLstat
      plainFs.statSync = origStat
      hook.deregister()
    },
  }
}

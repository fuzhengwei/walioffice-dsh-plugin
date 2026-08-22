import type { UserConfig } from 'tsdown'

const DEEPSEEK_EXTERNAL = /^@deepseek-ai\//
const WALIOFFICE_INTERNAL = /^@walioffice\//
const CLIENT_EXTERNALS = [
  'react',
  'react/jsx-runtime',
  'react-dom',
  'react-dom/client',
  '@deepseek-ai/cordis',
  '@deepseek-ai/dsh-client-runtime/client',
  '@deepseek-ai/dsh-client-ui-conversation/client',
  '@deepseek-ai/dsh-client-ui-layout/client',
  '@deepseek-ai/dsh-client-ui-slots',
  '@deepseek-ai/dsh-client-ui-tool/client',
] as const

export default [
  {
    name: 'walioffice-dsh-plugin',
    entry: ['src/index.ts'],
    outDir: 'lib',
    format: ['esm'],
    platform: 'node',
    target: 'es2024',
    dts: true,
    clean: true,
    fixedExtension: false,
    deps: {
      neverBundle: [DEEPSEEK_EXTERNAL, 'docx', 'exceljs', 'pptxgenjs'],
      alwaysBundle: [WALIOFFICE_INTERNAL],
      dts: {
        neverBundle: [DEEPSEEK_EXTERNAL, 'docx', 'exceljs', 'pptxgenjs'],
        alwaysBundle: [WALIOFFICE_INTERNAL],
      },
    },
  },
  {
    name: 'walioffice-dsh-plugin/client',
    entry: { client: 'src/client/index.ts' },
    outDir: 'lib',
    format: 'cjs',
    platform: 'browser',
    target: 'es2022',
    dts: false,
    sourcemap: true,
    clean: false,
    deps: {
      neverBundle: [...CLIENT_EXTERNALS],
    },
    outputOptions: {
      entryFileNames: 'client.js',
      banner: 'window.__ModuleLoader__.load({ id: "walioffice-dsh-plugin", factory: (require) => {',
      footer: 'return module.exports; } });',
      intro: 'var module = { exports: {} }; var exports = module.exports;',
    },
  },
] satisfies UserConfig[]

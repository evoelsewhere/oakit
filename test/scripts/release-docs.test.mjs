import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

describe('release documentation', () => {
  it('keeps README profile claims aligned with the producer matrix', async () => {
    const readme = await readFile(path.resolve('README.md'), 'utf8');
    const matrix = JSON.parse(
      await readFile(
        path.resolve('docs', 'evidence', '0.0.1', 'producer-matrix.json'),
        'utf8',
      ),
    );

    for (const [kind, profile] of Object.entries(matrix.profiles)) {
      expect(readme, kind).toContain(profile.profileId);
      expect(readme, kind).toContain(`effective \`${profile.level}\``);
    }
    expect(readme).toContain('do not claim arbitrary PPTX editing');
  });

  it('binds 0.0.1 evidence while preparing the CLI packaging hotfix', async () => {
    const changelog = await readFile(path.resolve('CHANGELOG.md'), 'utf8');
    const checklist = await readFile(
      path.resolve('docs', 'release-0.0.1.md'),
      'utf8',
    );
    const metadata = JSON.parse(await readFile(path.resolve('package.json')));
    const gates = JSON.parse(
      await readFile(
        path.resolve('docs', 'evidence', '0.0.1', 'release-gates.json'),
        'utf8',
      ),
    );

    expect(metadata.version).toBe('0.0.2');
    expect(changelog).toContain('## 0.0.2 - 2026-08-18');
    expect(changelog).toContain('## 0.0.1 - 2026-08-18');
    expect(changelog).not.toContain('## Unreleased');
    expect(checklist).toContain(
      '[run 32045412714](https://github.com/evoelsewhere/oakit/actions/runs/32045412714)',
    );
    expect(checklist).toContain(
      '[run 32049829830](https://github.com/evoelsewhere/oakit/actions/runs/32049829830)',
    );
    expect(checklist).toContain('21,184 mutants');
    expect(checklist).toContain('- [x] Pack and install the release tarball');
    expect(checklist).toContain(
      '[run 32053915220](https://github.com/evoelsewhere/oakit/actions/runs/32053915220)',
    );
    expect(gates).toMatchObject({
      ci: { conclusion: 'success', runId: '32049822932' },
      mutation: {
        compileError: 4450,
        killed: 16734,
        missed: 0,
        runId: '32049829830',
        total: 21184,
      },
      producer: {
        runId: '32045412714',
        templateCount: 30,
        totalSlides: 733,
      },
      releaseCandidate: {
        ci: { conclusion: 'success', jobs: 11, runId: '32053915220' },
        packageSmoke: {
          cjs: true,
          cliVersion: 'oakit 0.0.1',
          esm: true,
          officeFreePng: true,
          packedInstall: true,
          portableRestoreByteExact: true,
          subpathExports: true,
        },
        revision: '83f1cebf0dd1429a1f8305389681cc50e010b038',
        sourceTree: '4b7c749f58f908967b9868716bddd78e05a0fbcd',
      },
      release: {
        knownPackagingIssue: {
          apiExportsAffected: false,
          fixedIn: '0.0.2',
          scope: 'installed CLI mapping',
        },
        publish: {
          conclusion: 'success',
          provenance: true,
          runId: '32054316824',
        },
        revision: '42a7483fd101c2374cf6afe4e656f2da19787679',
        tag: 'v0.0.1',
      },
      sourceTree: '4b7c749f58f908967b9868716bddd78e05a0fbcd',
      version: '0.0.1',
    });
    expect(checklist).toContain('- [x] Create tag `v0.0.1`');
    expect(checklist).toContain('package smoke now executes');
    const hotfixChecklist = await readFile(
      path.resolve('docs', 'release-0.0.2.md'),
      'utf8',
    );
    expect(hotfixChecklist).toContain('node_modules/.bin/oakit');
    expect(hotfixChecklist).toContain(
      '[run 32054867081](https://github.com/evoelsewhere/oakit/actions/runs/32054867081)',
    );
  });
});

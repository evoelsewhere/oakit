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

  it('keeps the unreleased version blocked on final-revision evidence', async () => {
    const changelog = await readFile(path.resolve('CHANGELOG.md'), 'utf8');
    const checklist = await readFile(
      path.resolve('docs', 'release-0.0.1.md'),
      'utf8',
    );
    const metadata = JSON.parse(await readFile(path.resolve('package.json')));

    expect(metadata.version).toBe('0.0.0');
    expect(changelog).toContain('## Unreleased');
    expect(checklist).toContain(
      '- [ ] Refresh the controlled Google credential',
    );
    expect(checklist).toContain(
      '- [ ] Change `package.json` and lockfile metadata from `0.0.0` to `0.0.1`',
    );
  });
});

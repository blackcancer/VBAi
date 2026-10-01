"""Regression tests for the documentation checker, using disposable directories."""
import tempfile
import unittest
from pathlib import Path

from check_docs import anchors, check, prose


class DocumentationChecks(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def write(self, name, content):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding='utf-8')

    def test_valid_relative_link_and_anchor(self):
        self.write('README.md', '# Project\n[Guide](docs/start.md#first-step)\n')
        self.write('docs/start.md', '# Guide\n## First step\n')
        self.assertEqual(check(self.root), (2, 1, []))

    def test_missing_file(self):
        self.write('README.md', '# Project\n[Missing](missing.md)\n')
        self.assertIn('missing local destination', check(self.root)[2][0])

    def test_missing_anchor(self):
        self.write('README.md', '# Project\n[Missing](#gone)\n')
        self.assertIn('missing heading anchor', check(self.root)[2][0])

    def test_duplicate_heading_anchors(self):
        self.assertEqual(anchors('# A\n## A\n## A\n'), {'a', 'a-1', 'a-2'})

    def test_collision_with_explicit_numeric_heading(self):
        self.assertEqual(anchors('# A\n## A-1\n## A\n'), {'a', 'a-1', 'a-2'})

    def test_unicode_and_formatting_in_heading(self):
        self.assertEqual(anchors('# Café: `API` & **VBE**\n'), {'café-api--vbe'})

    def test_frontmatter_and_fenced_examples_are_ignored(self):
        self.write('README.md', '---\nname: template\n---\n# Project\n```md\n# Example\n[x](gone.md)\n```\n')
        self.assertEqual(check(self.root), (1, 0, []))

    def test_short_fence_does_not_close_long_fence(self):
        self.assertNotIn('hidden', prose('# A\n````md\n```\nhidden\n````\n'))

    def test_comments_are_ignored(self):
        self.write('README.md', '# Project\n<!-- [x](missing.md) -->\n')
        self.assertEqual(check(self.root), (1, 0, []))

    def test_external_urls_are_not_fetched(self):
        self.write('README.md', '# Project\n[x](https://example.invalid/no-file)\n[x](mailto:team@example.invalid)\n')
        self.assertEqual(check(self.root), (1, 0, []))

    def test_html_image_is_checked(self):
        self.write('README.md', '# Project\n<img src="missing.png" alt="Logo">\n')
        self.assertIn('missing local destination', check(self.root)[2][0])

    def test_markdown_image_is_checked(self):
        self.write('README.md', '# Project\n![Logo](missing.png)\n')
        self.assertIn('missing local destination', check(self.root)[2][0])

    def test_percent_encoded_path_and_anchor(self):
        self.write('README.md', '# Project\n[x](other%20page.md#caf%C3%A9)\n')
        self.write('other page.md', '# Café\n')
        self.assertEqual(check(self.root), (2, 1, []))

    def test_link_cannot_escape_root(self):
        self.write('README.md', '# Project\n[x](../outside.md)\n')
        self.assertIn('escapes the repository', check(self.root)[2][0])

    def test_sparse_verified_asset_path(self):
        self.write('README.md', '# Project\n![Logo](assets/logo.png)\n')
        self.assertEqual(check(self.root, {'assets/logo.png'}), (1, 1, []))

    def test_sparse_manifest_does_not_invent_anchor_contents(self):
        self.write('README.md', '# Project\n[x](missing.md#title)\n')
        self.assertIn('cannot inspect Markdown anchor', check(self.root, {'missing.md'})[2][0])

    def test_test_prompt_fixtures_are_preserved_outside_editorial_scope(self):
        self.write('README.md', '# Project\n')
        self.write('tests/VBAi.Tests/Infrastructure/Fixtures/Prompts/input.md', 'not documentation')
        self.write('tests/VBAi.Tests/Infrastructure/Fixtures/Editor/matrix.md', 'not documentation')
        self.write('tests/Infrastructure/.agents/skills/example/SKILL.md', 'not documentation')
        self.assertEqual(check(self.root), (1, 0, []))

    def test_generated_artifacts_are_excluded(self):
        self.write('README.md', '# Project\n')
        self.write('artifacts/test.md', 'unstructured output')
        self.assertEqual(check(self.root), (1, 0, []))

    def test_isolated_worktrees_do_not_duplicate_editorial_scope(self):
        self.write('README.md', '# Project\n[Guide](docs/guide.md)\n')
        self.write('docs/guide.md', '# Guide\n')
        self.write('.worktrees/parallel/docs/guide.md', 'unfinished parallel draft')
        self.write('.worktrees/parallel/tests/VBAi.Tests/Infrastructure/Fixtures/input.md', 'not documentation')
        self.assertEqual(check(self.root), (2, 1, []))

    def test_multiple_h1_is_error(self):
        self.write('README.md', '# First\n# Second\n')
        self.assertIn('exactly one', check(self.root)[2][0])

    def test_final_newline(self):
        self.write('README.md', '# Project')
        self.assertIn('final newline', check(self.root)[2][0])


if __name__ == '__main__':
    unittest.main()

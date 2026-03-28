from __future__ import annotations

import re
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent
README_PATH = REPO_ROOT / 'README.md'


HEADING_PATTERN = re.compile(r'^(#{2,6})\s+(.+?)\s*$', re.MULTILINE)


def normalize_heading(text: str) -> str:
    return ' '.join(text.split()).casefold()


def get_section(content: str, heading: str) -> str:
    matches = list(HEADING_PATTERN.finditer(content))
    target_heading = normalize_heading(heading)

    for index, match in enumerate(matches):
        if normalize_heading(match.group(2)) != target_heading:
            continue

        current_level = len(match.group(1))
        section_start = match.end()
        section_end = len(content)

        for later_match in matches[index + 1:]:
            if len(later_match.group(1)) <= current_level:
                section_end = later_match.start()
                break

        return content[section_start:section_end].strip()

    return ''


def section_index(content: str, heading: str) -> int:
    matches = list(HEADING_PATTERN.finditer(content))
    target_heading = normalize_heading(heading)

    for match in matches:
        if normalize_heading(match.group(2)) == target_heading:
            return match.start()

    return -1


def find_heading_index(content: str, headings: list[str]) -> int:
    matches = list(HEADING_PATTERN.finditer(content))
    normalized_targets = {normalize_heading(heading) for heading in headings}

    for match in matches:
        if normalize_heading(match.group(2)) in normalized_targets:
            return match.start()

    return -1


def assert_contains_any(test_case: unittest.TestCase, text: str, options: list[str], message: str) -> None:
    normalized_text = text.casefold()
    if any(option.casefold() in normalized_text for option in options):
        return
    test_case.fail(message)


def assert_regex_any(
    test_case: unittest.TestCase,
    text: str,
    patterns: list[str],
    message: str,
    flags: int = 0,
) -> None:
    for pattern in patterns:
        if re.search(pattern, text, flags):
            return
    test_case.fail(message)


def assert_in_order(test_case: unittest.TestCase, text: str, snippets: list[str], message: str) -> None:
    position = -1
    lowered_text = text.casefold()

    for snippet in snippets:
        next_position = lowered_text.find(snippet.casefold(), position + 1)
        if next_position == -1:
            test_case.fail(message)
        position = next_position


class ReadmeContentTests(unittest.TestCase):
    def test_readme_includes_bilingual_user_facing_usage_guidance(self):
        content = README_PATH.read_text(encoding='utf-8')
        english_section = get_section(content, 'English')
        traditional_chinese_section = get_section(content, '繁體中文')
        english_index = section_index(content, 'English')
        traditional_chinese_index = section_index(content, '繁體中文')

        self.assertTrue(english_section, 'README.md must include an English markdown section.')
        self.assertTrue(
            traditional_chinese_section,
            'README.md must include a 繁體中文 markdown section.',
        )
        self.assertLess(
            english_index,
            traditional_chinese_index,
            'README.md must present the English section before the 繁體中文 section.',
        )
        assert_regex_any(
            self,
            content,
            [
                r'original README.*repo title',
                r'provenance.*repo title',
                r'最初.*README.*repo title',
            ],
            'README.md must include a provenance note explaining that the original README only had the repo title.',
            flags=re.IGNORECASE | re.DOTALL,
        )

        english_required_snippets = [
            'workbook',
            'verify',
            '四號谷地_視覺版',
            '四號谷地_表格版',
            '四號谷地_1',
            'pwsh -File .\\scripts\\apply_sihao_valley_table.ps1 -EnableWorkbookWrites',
            '-WorkbookPath',
            'python .\\scripts\\verify_sihao_valley_table.py',
        ]
        traditional_chinese_required_snippets = [
            '搜尋文字',
            '驗證',
            '工作表',
            '四號谷地_視覺版',
            '四號谷地_表格版',
            '四號谷地_1',
            'pwsh -File .\\scripts\\apply_sihao_valley_table.ps1 -EnableWorkbookWrites',
            '-WorkbookPath',
            'python .\\scripts\\verify_sihao_valley_table.py',
        ]

        for snippet in english_required_snippets:
            with self.subTest(section='English', snippet=snippet):
                self.assertIn(snippet, english_section)

        for snippet in traditional_chinese_required_snippets:
            with self.subTest(section='繁體中文', snippet=snippet):
                self.assertIn(snippet, traditional_chinese_section)

        assert_in_order(
            self,
            english_section,
            ['project overview', 'workbook overview', 'which sheet to use', 'filter and search', 'generated vs source', 'automation', 'provenance'],
            'English section must follow the requested topic flow.',
        )
        assert_in_order(
            self,
            traditional_chinese_section,
            ['專案概覽', '活頁簿概覽', '使用哪個工作表', '篩選與搜尋', '產生內容與來源', '自動化', '來源說明'],
            '繁體中文 section must follow the requested topic flow.',
        )

        english_project_overview_index = find_heading_index(english_section, ['Project overview'])
        english_overview_index = find_heading_index(english_section, ['Workbook overview'])
        english_sheet_usage_index = find_heading_index(english_section, ['Which sheet to use'])
        traditional_chinese_project_overview_index = find_heading_index(traditional_chinese_section, ['專案概覽'])
        traditional_chinese_overview_index = find_heading_index(traditional_chinese_section, ['活頁簿概覽'])
        traditional_chinese_sheet_usage_index = find_heading_index(traditional_chinese_section, ['使用哪個工作表'])

        self.assertNotEqual(
            english_project_overview_index,
            -1,
            'English section must include a distinct project overview heading.',
        )
        self.assertNotEqual(
            english_overview_index,
            -1,
            'English section must include a distinct workbook overview heading.',
        )
        self.assertNotEqual(
            traditional_chinese_project_overview_index,
            -1,
            '繁體中文 section must include a distinct project overview heading.',
        )
        self.assertNotEqual(
            traditional_chinese_overview_index,
            -1,
            '繁體中文 section must include a distinct workbook overview heading.',
        )
        self.assertLess(
            english_project_overview_index,
            english_overview_index,
            'English project overview heading must appear before workbook overview.',
        )
        self.assertLess(
            english_overview_index,
            english_sheet_usage_index,
            'English workbook overview heading must appear before sheet-usage guidance.',
        )
        self.assertLess(
            traditional_chinese_project_overview_index,
            traditional_chinese_overview_index,
            '繁體中文 project overview heading must appear before workbook overview.',
        )
        self.assertLess(
            traditional_chinese_overview_index,
            traditional_chinese_sheet_usage_index,
            '繁體中文 workbook overview heading must appear before sheet-usage guidance.',
        )

        assert_regex_any(
            self,
            english_section,
            [r'搜尋文字.*?(filter|search)', r'(filter|search).*?搜尋文字'],
            'English section must explain how `搜尋文字` is used.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            traditional_chinese_section,
            [r'搜尋文字.*?(篩選|查找|搜尋)', r'(篩選|查找|搜尋).*?搜尋文字'],
            '繁體中文 section must explain how `搜尋文字` is used.',
            flags=re.DOTALL,
        )
        assert_regex_any(
            self,
            english_section,
            [
                r'四號谷地_視覺版.*(primary|normal).*四號谷地_表格版',
                r'四號谷地_表格版.*(primary|normal).*四號谷地_視覺版',
            ],
            'English section must describe `四號谷地_視覺版` and `四號谷地_表格版` as the main user-facing workflow.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            english_section,
            [r'四號谷地_1.*(source|backend)'],
            'English section must describe `四號谷地_1` as a source/backend sheet.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            traditional_chinese_section,
            [
                r'四號谷地_視覺版.*(主要|一般使用).*四號谷地_表格版',
                r'四號谷地_表格版.*(主要|一般使用).*四號谷地_視覺版',
            ],
            '繁體中文 section must describe `四號谷地_視覺版` and `四號谷地_表格版` as the main user-facing workflow.',
            flags=re.DOTALL,
        )
        assert_regex_any(
            self,
            traditional_chinese_section,
            [r'四號谷地_1.*(來源|後端)'],
            '繁體中文 section must describe `四號谷地_1` as a source/backend sheet.',
            flags=re.DOTALL,
        )

        assert_regex_any(
            self,
            english_section,
            [
                r'四號谷地_表格版.*slicer',
                r'四號谷地_表格版.*built-in filter',
                r'slicer.*四號谷地_表格版',
                r'filter.*四號谷地_表格版',
            ],
            'English section must explain slicers and built-in filters on `四號谷地_表格版`.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            traditional_chinese_section,
            [
                r'四號谷地_表格版.*slicer',
                r'四號谷地_表格版.*篩選',
                r'slicer.*四號谷地_表格版',
                r'篩選.*四號谷地_表格版',
            ],
            '繁體中文 section must explain slicers and built-in filters on `四號谷地_表格版`.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            english_section,
            [r'(generated|managed).*(manual entry|manually enter).*(rebuild)', r'rebuild.*(generated|managed).*(manual entry|manually enter)'],
            'English section must warn that generated structures are not manual-entry areas unless the user plans to rebuild them again.',
            flags=re.IGNORECASE | re.DOTALL,
        )
        assert_regex_any(
            self,
            traditional_chinese_section,
            [r'(產生|受控).*(手動輸入|手動填寫).*(重建)', r'重建.*(產生|受控).*(手動輸入|手動填寫)'],
            '繁體中文 section must warn that generated structures are not manual-entry areas unless the user plans to rebuild them again.',
            flags=re.DOTALL,
        )


if __name__ == '__main__':
    unittest.main()

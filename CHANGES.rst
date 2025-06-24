1.1.2 (24-06-2025)
--------------------
- implemented PDF file content sanity-check (PowerPointReport._convert_pdf)
- fixed numpy import
- replace pkg_resources (deprecated) with importlib_resources
- removed Pillow restriction as it caused an installation error and the new version is now supported by python-pptx (>= v0.6.23)

1.1.1 (03-07-2023)
--------------------
- Fixed bug where missing files in `grouped_content` raised an error even when missing_file="text". Combinations of empty_slide=="skip" and missing_file="text"/"empty"/"skip" now work the same as for regular content lists.

1.1.0 (23-06-2023)
--------------------
- Reworked markdown parsing to allow for headers (`# header`), lists (`- item 1`) and nested types (`**partly bold and _italics_ string**`). Unsupported markdown types are logged as warnings and added as plain text.
- Content patterns with invalid regexes, e.g. "**text**" no longer raise an error, but are now logged as warnings and added as text (which may or may not contain markdown).
- Added restriction of pillow<10 due to python-pptx deprecation: "DeprecationWarning: getsize is deprecated and will be removed in Pillow 10 (2023-07-01). Use getbbox or getlength instead."

1.0.0 (02-06-2023)
--------------------
- Initial release

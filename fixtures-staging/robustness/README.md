# robustness/ — loader crash-test fixtures (NOT parity)

These four .docx are deliberately damaged. They are for the e2e loader
(does WordInWeb degrade gracefully rather than crash?), not for pixel
parity — Word will refuse or repair them, so there is no ground-truth PDF.

- `damaged-missing-styles.docx` — document.xml.rels points at styles.xml but the part is absent.
- `damaged-unknown-elements.docx` — unknown elements in the w: namespace and a foreign namespace sprinkled through the body.
- `damaged-2mb-paragraph.docx` — one paragraph with a single ~2MB text run (perf / memory).
- `damaged-truncated.docx` — a valid package with its last 400 bytes chopped (broken central directory).

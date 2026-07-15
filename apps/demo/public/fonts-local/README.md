# fonts-local (development only)

This directory holds the real Microsoft Office fonts used to make the demo's
canvas measurement and glyph rendering match Word exactly. The font binaries
are licensed. Don't deploy your commercial apps with this you fool.

## Graceful fallback

If the local files are absent, the browser falls back to the bundled
metric-compatible substitutes. The viewer still works, but glyph appearance
and line wrapping can differ from Microsoft Word.

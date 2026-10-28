"""Weekly support-report generation pipeline.

Discover (input manifest), parse (Part A phone logs + CarlSource -> bundle.json), and
build (Part A .docx) stages are all implemented. See `reportgen.run` for the CLI, and
`reportgen.discover` / `reportgen.bundle` / `reportgen.build_part_a` for the stages.
"""

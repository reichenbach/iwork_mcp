---
name: keynote-deck
description: Build a polished Keynote presentation with iwork-mcp tools.
---

# Build a Keynote deck

Use `keynote_list_themes` and `keynote_list_master_slides` when the user has not specified a visual direction. Prefer `keynote_create_presentation_with_slides` to create the initial deck efficiently. Use `keynote_set_slide_title`, `keynote_set_slide_body`, and `keynote_set_presenter_notes` for revisions.

Keep each slide focused on one idea. Use short titles, concise body text, and presenter notes for detail. Add supporting visuals with `keynote_add_image_to_slide`; use `keynote_position_item`, `keynote_align_items`, and `keynote_distribute_items` for consistent layout. Use `keynote_add_table_to_slide` only when a table communicates the data more clearly than prose.

Before finishing, inspect the result with `keynote_list_slides`, `keynote_get_slide_content`, and `keynote_list_slide_items`. Save with `keynote_save_presentation` and export with `keynote_export_presentation` only when requested.

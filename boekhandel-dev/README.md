# Velles i Novelles Apps Script (DEV)

## ItemTypes sheet schema

The current DEV ItemTypes tab uses these columns in this order:

| Column | Purpose |
| --- | --- |
| item_type_id | Stable key stored on items; do not rename or change existing IDs. |
| name_ca | Catalan display name. |
| active | Boolean flag (TRUE to show the type, FALSE to hide it). |
| name_es | Spanish display name. |
| name_en | English display name. |

The new-item and inventory-filter dropdowns display the name for the configured APP_LANGUAGE (ca, es, or en). If that translation is blank, the app falls back to name_ca, then to item_type_id. When adding an item type, keep its ID stable and fill in all three names. Existing item IDs and active flags should be preserved during sheet migrations.

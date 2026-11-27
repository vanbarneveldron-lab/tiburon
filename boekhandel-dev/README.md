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


## Sales module (DEV)

A sale is created from an active inventory item. Each sale stores the item ID and reference, ISO sale date, required free-text sales channel, VAT-inclusive gross price, VAT rate, VAT amount, net amount, invoice placeholders, and status in the existing `Sales` sheet schema. The VAT rate defaults to 4% in the form and accepts a rate from 0% to 100% with up to two decimal places. Sales can be reversed: the sale row remains for the audit trail with status `reversed`, and the item returns to `active`. Reversed sales are retained in history and excluded from the active sales count. The channel is free text until the publishing module can provide platform choices.

The Publications screen records manual listings against active inventory items and configured active Platforms, with listing and withdrawal dates. Withdrawal keeps the record for history. The Reports screen summarizes active sales between optional inclusive start and end dates (defaulting to the current year through today), including gross, VAT, net, and count. Reversed sales are excluded.

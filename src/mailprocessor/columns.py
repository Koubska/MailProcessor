"""Names of the data sheet columns the app fills itself. Field rules must not use them (see `config.Profile`)."""

# After the field columns, in this order.
RECEIVED_COLUMN = "Eingegangen am"  # the mail's Date header
TRANSFERRED_COLUMN = "Übertragen am"  # when the run added the row
CONTENT_COLUMN = "E-Mail-Inhalt"  # the full mail text
FIXED_DATA_COLUMNS = (RECEIVED_COLUMN, TRANSFERRED_COLUMN, CONTENT_COLUMN)
# First column of the shared data sheet when there are several profiles: which profile the row was read with.
PROFILE_COLUMN = "Profil"

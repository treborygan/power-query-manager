"""Small helpers for dialogs that should stay within their owning workspace."""


def resize_dialog_to_parent(dialog, default_width: int, default_height: int) -> None:
    """Use the parent workspace as the initial dialog size limit when available.

    Keep the dialog's declared minimum size intact. Top-level dialogs without a
    parent retain their existing design dimensions.
    """
    parent = dialog.parentWidget()
    if parent is None:
        dialog.resize(default_width, default_height)
        return

    width = max(dialog.minimumWidth(), min(default_width, parent.width()))
    height = max(dialog.minimumHeight(), min(default_height, parent.height()))
    dialog.resize(width, height)

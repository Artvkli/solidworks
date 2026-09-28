def com_value(obj, name, default=None):
    """
    Read a COM member (property or method) safely.

    In pywin32, some SolidWorks members come back as a value and some
    as a callable method. This returns the real value in both cases.
    COM objects themselves (CDispatch) are NOT called.
    """

    try:
        value = getattr(obj, name)
    except Exception:
        return default

    # COM objects have _oleobj_ -> they are values, not methods
    if callable(value) and not hasattr(value, "_oleobj_"):
        try:
            return value()
        except Exception:
            return default

    return value

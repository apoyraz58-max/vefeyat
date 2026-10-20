import re
import unicodedata


def normalize_name(text):
    """Search key only. Never replace displayed names with this value."""
    text = str(text or "").translate(str.maketrans({"İ": "i", "ı": "i"})).casefold()
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    words = re.sub(r"[^\w]+", " ", text, flags=re.UNICODE).replace("_", " ").split()
    return " ".join("ibn" if word == "ibni" else word for word in words)

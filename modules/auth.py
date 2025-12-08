# modules/auth.py
import hashlib

SALT = "viettel_secure_salt_2025"  # nếu muốn đổi, để trong st.secrets

def hash_password(plain: str, salt: str = SALT) -> str:
    return hashlib.sha256((salt + plain).encode("utf-8")).hexdigest()

def verify_password(provided: str, stored_hash: str, salt: str = SALT) -> bool:
    return hash_password(provided, salt) == stored_hash

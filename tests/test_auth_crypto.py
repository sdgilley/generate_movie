from webapp.auth import encrypt_secret, decrypt_secret


def test_encrypt_decrypt_roundtrip():
    secret = 'this-is-a-secret'
    enc = encrypt_secret(secret)
    assert enc != secret
    dec = decrypt_secret(enc)
    assert dec == secret

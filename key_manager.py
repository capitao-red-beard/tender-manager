import keyring
import keyring.errors


def set_password(service_name, username, password):
    try:
        keyring.set_password(service_name, username, password)
    except keyring.errors.PasswordSetError as k:
        print(str(k))


def get_password(service_name, username):
    try:
        return keyring.get_password(service_name, username)

    except keyring.errors.KeyringError as k:
        print(str(k))


def get_credentials(service_name, username):
    try:
        return keyring.get_credential(service_name, username)

    except keyring.errors.KeyringError as k:
        print(str(k))


def delete_password(service_name, username):
    try:
        keyring.delete_password(service_name, username)
    except keyring.errors.PasswordDeleteError as k:
        print(str(k))

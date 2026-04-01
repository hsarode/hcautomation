import rsa
import getpass
from pathlib import Path
from importlib import resources

def get_app_dir() -> Path:
    app_dir = Path.home() / ".hcautomation"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir

def generate_login_file():
    with resources.files("hcautomation").joinpath("pub_key.PEM").open("rb") as f:
        pub_key = rsa.PublicKey.load_pkcs1(f.read())

    er_usrname = str(input('Please enter ER username:'))
    er_pswd = getpass.getpass('Enter ER password')
    er = rsa.encrypt(f"{er_usrname}###{er_pswd}".encode('utf-8'), pub_key)
    er_login = get_app_dir() / 'er_login.txt'
    er_login.write_bytes(er)

if __name__ == "__main__":
    generate_login_file()
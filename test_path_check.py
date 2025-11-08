import os
import glob

path = r"\\PC011\Users\yasumoku\Desktop\タカラ関係\工程表"
pattern = os.path.join(path, "10-25*.xls")

print("検索パターン:", pattern)
files = glob.glob(pattern)
print("見つかったファイル一覧:", files)
for f in files:
    print(f, "→", os.path.exists(f))

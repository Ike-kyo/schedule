from schedule_logic import DEFAULT_IMPORT, DEFAULT_OUTPUT
import os

print("DEFAULT_IMPORT =", DEFAULT_IMPORT)
print("存在するか？", os.path.exists(DEFAULT_IMPORT))
print("DEFAULT_IMPORTの内容:", DEFAULT_IMPORT)
print("os.listdir(DEFAULT_IMPORT):", os.listdir(DEFAULT_IMPORT))

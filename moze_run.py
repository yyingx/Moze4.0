import glob, subprocess, sys

scripts = sorted(glob.glob(r"e:\天之逸2025\Moze4.0\Moze4.0_Import_v*.py"))
if scripts:
    latest = scripts[-1]
    print(f"启动版本: {latest.split('_v')[-1].replace('.py', '')}")
    subprocess.run([sys.executable, latest] + sys.argv[1:])
else:
    input("未找到 Moze4.0_Import_v*.py，按 Enter 退出")

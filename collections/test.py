import json
import requests

usd_user = requests.get(
    "https://proton.cryptolions.io/v2/state/get_tokens?limit=1000&account=pimlrp"
).text
usd_user = json.loads(usd_user)
for x in usd_user["tokens"]:
    if x["symbol"] == "XUSDC":
        USDC = x["amount"]
upper5 = 0.035 * USDC
num6_20 = 0.025 * USDC
num21_40 = 0.0125 * USDC
print(USDC)
print(str(upper5) + " " + str(num6_20) + " " + str(num21_40))

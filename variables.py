import json
import os

from utils import resolve_relative_path

save_files_to_debug = True


def get_default_values(option):
    if option == 0:  # DAX
        return dict(
            Spannweite=2000,
            step=50,
            volatility_Laufzeit=60,
            contract_value=5,
        )

    else:  # STOXX
        return dict(
            Spannweite=700,
            step=25,
            volatility_Laufzeit=365,
            contract_value=1,
        )


dict_index_stock = {
    0: "DAX",
    1: "STOXX"
}

dict_stock_index = {v: k for k, v in dict_index_stock.items()}

# Folders where results will be stored
current_results_path = resolve_relative_path("./Results")
temp_results_path = resolve_relative_path("./Results/temp")
old_results_path = resolve_relative_path("./Results/Old")

# Source file folder
src_path = "./src/"

# Source files
src_html = os.path.join(src_path, "template.html")
src_css = os.path.join(src_path, "style.css")

email_messages_path = os.path.join(src_path, "email_messages")

list_emails_path = os.path.join(src_path, "list_emails.json")
with open(list_emails_path) as file:
    list_emails = json.load(file)['emails']


summery_html = os.path.join(temp_results_path, "summery.html")
summery_expiry_html = os.path.join(temp_results_path, "summery_expiry.html")
summery_basic_html = os.path.join(temp_results_path, "summery_basic.html")

result_css = os.path.join(temp_results_path, "style.css")
result_image = os.path.join(temp_results_path, "image.svg")

default_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36",
    "Accept-Encoding": "gzip, deflate",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
    "image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "DNT": "1",
    "Connection": "close",
    "Upgrade-Insecure-Requests": "1",
    "Pragma": "no-cache",
    "Cache-Control": "no-cache",
    "SEC-CH-UA": "\" Not A;Brand\";v=\"99\", \"Chromium\";v=\"90\", \"Google Chrome\";v=\"90\"",
    "SEC-CH-UA-ARCH": "\"x86\"",
    "SEC-CH-UA-BROWSER": "\"Chromium\"",
    "SEC-CH-UA-MODE": "\"Regular\"",
    "SEC-CH-UA-PLATFORM": "\"Windows\"",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
}

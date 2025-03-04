import subprocess
from concurrent.futures import ThreadPoolExecutor

def run_script(script_name):
    subprocess.run(["python", script_name])

scripts = [
    "generate_url_for_keyword.py",
    "mercari_min_price_scraper.py",
]

with ThreadPoolExecutor() as executor:
    executor.map(run_script, scripts)



subprocess.run(["python", "combined_min_price_scraper.py"])
subprocess.run(["python", "ebay.py"])
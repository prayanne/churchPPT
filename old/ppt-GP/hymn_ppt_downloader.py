
import aiohttp
import asyncio
import re
import os
from bs4 import BeautifulSoup

BASE_URL = "https://godpeople.or.kr/hymn/category/3186073"
DOWNLOAD_DIR = "ppt_downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

headers = {"User-Agent": "Mozilla/5.0"}

def extract_download_url(html: str):
    match = re.search(r'window\.location\s*=\s*[\'"](/index\.php\?act=procFileDownload[^\'"]+)', html)
    if match:
        return "https://godpeople.or.kr" + match.group(1)
    return None

async def fetch_html(session, url):
    async with session.get(url, headers=headers) as response:
        return await response.text()

async def fetch_binary(session, url):
    async with session.get(url, headers=headers) as response:
        return await response.read()

async def get_detail_links_from_page(session, page_num):
    page_url = f"{BASE_URL}?page={page_num}"
    html = await fetch_html(session, page_url)
    soup = BeautifulSoup(html, "html.parser")
    return [
        "https://godpeople.or.kr" + a['href']
        for a in soup.select("a.thumb-link") if a.get("href")
    ]

async def download_ppt_from_detail(session, detail_url):
    html = await fetch_html(session, detail_url)
    download_url = extract_download_url(html)
    if not download_url:
        print(f"[!] 다운로드 링크 없음: {detail_url}")
        return None

    ppt_data = await fetch_binary(session, download_url)
    title_soup = BeautifulSoup(html, "html.parser")
    title_tag = title_soup.select_one("div.hymn-title strong")
    title = title_tag.text.strip().replace(" ", "_") if title_tag else "unknown"
    title = re.sub(r'[\\/*?:"<>|]', "", title)

    filename = os.path.join(DOWNLOAD_DIR, f"{title}.pptx")
    with open(filename, "wb") as f:
        f.write(ppt_data)
    print(f"[✔] 다운로드 완료: {filename}")
    return filename

async def main():
    downloaded_files = []
    async with aiohttp.ClientSession() as session:
        for page in range(1, 34):
            print(f"[*] 페이지 {page} 처리 중...")
            detail_links = await get_detail_links_from_page(session, page)
            for link in detail_links:
                try:
                    filepath = await download_ppt_from_detail(session, link)
                    if filepath:
                        downloaded_files.append(filepath)
                except Exception as e:
                    print(f"[!] 오류 발생: {e}")
                await asyncio.sleep(1)
    print(f"총 {len(downloaded_files)}개 파일 다운로드 완료.")
    return downloaded_files

if __name__ == "__main__":
    import nest_asyncio
    nest_asyncio.apply()
    asyncio.get_event_loop().run_until_complete(main())

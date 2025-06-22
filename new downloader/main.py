import aiohttp
import asyncio
import re
import os
from bs4 import BeautifulSoup

BASE_URL = "https://godpeople.or.kr/hymn/category/3186073"
DOWNLOAD_DIR = "ppt_downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

headers = {"User-Agent": "Mozilla/5.0"}

async def fetch_html(session, url):
    print(f"[HTML 요청] {url}")
    async with session.get(url, headers=headers) as response:
        return await response.text()

async def fetch_binary(session, url, cookies=None):
    print(f"[파일 요청] {url}")
    async with session.get(url, headers=headers, cookies=cookies) as response:
        return await response.read()

async def get_post_links_from_page(session, page_num):
    page_url = f"{BASE_URL}?page={page_num}"
    print(f"[페이지 탐색] {page_url}")
    html = await fetch_html(session, page_url)
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.select("a.hx"):
        href = a.get("href")
        if href and href.startswith("/hymn/"):
            full_link = "https://godpeople.or.kr" + href
            print(f"  [게시글 발견] {full_link}")
            links.append(full_link)
    return links

async def get_download_page_url(session, post_url):
    print(f"[게시글 진입] {post_url}")
    html = await fetch_html(session, post_url)
    soup = BeautifulSoup(html, "html.parser")
    try:
        a_tag = soup.select("table.bd_tb")[1].select("li a")
        if a_tag[0].contents[0].endswith("pptx"):
            a_tag = a_tag[0]
        else:
            a_tag = a_tag[1]
    except (IndexError, AttributeError):
        print("  [!] 다운로드 링크 요소를 찾을 수 없음")
        return None

    href = a_tag.get("href")
    if href:
        download_url = "https://godpeople.or.kr" + href
        print(f"  [다운로드 페이지 링크 추출] {download_url}")
        return download_url
    print("  [!] 다운로드 페이지 링크 없음")
    return None

async def extract_real_download_url(session, download_page_url):
    if "procFileDownload" in download_page_url:
        print(f"  [바로 다운로드 가능한 링크 확인됨] {download_page_url}")
        return download_page_url

    html = await fetch_html(session, download_page_url)
    match = re.search(r'window\.location\s*=\s*["\'](/index\.php\?act=procFileDownload[^"\']+)', html)
    if match:
        real_link = "https://godpeople.or.kr" + match.group(1)
        print(f"  [다운로드 링크 추출됨 from script] {real_link}")
        return real_link

    print("  [!] 다운로드 링크 추출 실패")
    return None

async def download_ppt_from_download_page(session, download_page_url):
    print(f"[다운로드 페이지 진입] {download_page_url}")
    real_url = await extract_real_download_url(session, download_page_url)
    if not real_url:
        print(f"  [!] 다운로드 링크 없음: {download_page_url}")
        return None

    file_id_match = re.search(r'file_srl=(\d+)', real_url)
    file_id = file_id_match.group(1) if file_id_match else "unknown"
    cookie_key = f"download_ad_ok_{file_id}"
    cookies = {cookie_key: "true"}

    ppt_data = await fetch_binary(session, real_url, cookies=cookies)

    html = await fetch_html(session, download_page_url)
    soup = BeautifulSoup(html, "html.parser")
    title_tag = soup.select_one("head > title")
    title_text = title_tag.text.strip() if title_tag else "unknown"

    match = re.search(r'(\d+)', title_text)
    number_only = match.group(1) if match else "unknown"

    filename = os.path.join(DOWNLOAD_DIR, f"[새찬송가] {number_only}장.pptx")
    with open(filename, "wb") as f:
        f.write(ppt_data)
    print(f"[✔] 다운로드 완료: {filename}")
    return filename

async def handle_post_link(session, post_link, semaphore):
    async with semaphore:
        try:
            download_page_url = await get_download_page_url(session, post_link)
            if download_page_url:
                return await download_ppt_from_download_page(session, download_page_url)
        except Exception as e:
            print(f"[!] 오류 발생: {e}")
        return None

async def main():
    downloaded_files = []
    concurrency = 10
    semaphore = asyncio.Semaphore(concurrency)  # 동시에 최대 5개 다운로드 허용

    async with aiohttp.ClientSession() as session:
        for page in range(1, 34):
            print(f"\n======================\n[📄 페이지 {page} 처리 시작]\n======================")
            post_links = await get_post_links_from_page(session, page)
            tasks = [
                handle_post_link(session, link, semaphore)
                for link in post_links
            ]
            results = await asyncio.gather(*tasks)
            downloaded_files.extend([r for r in results if r])

    print(f"\n총 {len(downloaded_files)}개 파일 다운로드 완료.")

if __name__ == "__main__":
    import nest_asyncio
    nest_asyncio.apply()
    asyncio.get_event_loop().run_until_complete(main())

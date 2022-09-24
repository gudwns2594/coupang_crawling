import requests
from bs4 import BeautifulSoup
import pyautogui
import openpyxl

keyword = pyautogui.prompt("검색어를 입력하세요")

#엑셀 만들기
wb = openpyxl.Workbook("coupang_result.xlsx")
ws = wb.create_sheet(keyword)
ws.append(['순위','브랜드명','상품명','가격','상세페이지링크'])





rank = 1
done = False
for page in range(1,5):
    if done == True:
        break
    print(page,"번째 페이지 입니다.")
    url = f"https://www.coupang.com/np/search?component=&q={keyword}&page={page}"
    # 쿠팡홈페이지 접속하려면 밑에 코드가 필요함
    header = {
        'Host': 'www.coupang.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.8,en-US;q=0.5,en;q=0.3',
    }

    response = requests.get(url,headers=header)
    html = response.text
    soup = BeautifulSoup(html,"html.parser")
    links = soup.select("a.search-product-link")

    # 상세페이지 친구들
    for link in links:
        # 광고 상품 제거
        if len(link.select("span.ad-badge-text")) >0 :
            print("광고상품 입니다.")
        else: 
            sub_url = "https://www.coupang.com/" + link.attrs['href']
            response = requests.get(sub_url,headers=header)
            html = response.text
            soup = BeautifulSoup(html,"html.parser")
            
            
            # select 함수는 리스트로 반환한다. (주의!!)
            #브랜드명 (있을수도 있고 없을수도 있음)
            # -중고상품일 때는 태그가 달라짐
            # try -except 로 예외처리 해줌
            try:
                brand_name = soup.select_one("a.prod-brand-name").text
            except:
                brand_name = soup.select_one("a.prod-brand-name").text
            #상풍명
            brand_name = soup.select_one("a.prod-brand-name").text
            #브랜드명    
            goods_name = soup.select_one("h2.prod-buy-header__title").text
            #가격
            try:
                price = soup.select_one("span.total-price > strong").text
            except:
                price = soup.select_one("span.total-price > strong").text
            price = soup.select_one("span.total-price > strong").text
            print(rank,brand_name,goods_name,price)
            ws.append([rank,brand_name,goods_name,price,sub_url])
            rank += 1
            if rank > 100:
                done = True   # 가장 가까운 반복문 탈출 시킴
                break
            
wb.save('coupang_result.xlsx')
    
    
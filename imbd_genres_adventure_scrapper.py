from bs4 import BeautifulSoup
import requests, openpyxl, re

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Top 250 Movies"
sheet.append(["Rank", "Movie Name", "Director", "Rating", "Year of release", "Story scinapse", "Gross"])

try:
    response = requests.get("https://www.imdb.com/search/title/?genres=adventure&sort=user_rating,desc&title_type=feature&num_votes=25000,&pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=5aab685f-35eb-40f3-95f7-c53f09d542c3&pf_rd_r=AC6MPTDD7DA3CVW9KHHZ&pf_rd_s=right-6&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_gnr_2")
    soup = BeautifulSoup(response.text, "html.parser")
    movie = soup.find("div", class_="lister-list").find_all("div", class_="lister-item mode-advanced")

    for movies in movie:
        rank = movies.find("span", class_="lister-item-index").text[0]
        movie_name = movies.find("h3", class_="lister-item-header").a.text
        # director = movies.
        rating = movies.strong.text
        year = movies.find("span", class_ ="lister-item-year text-muted unbold").text
        year = re.sub("D","",year)
        story = movies.find("p").findNext("p").get_text(strip=True)
        director = movies.find("p").findNext("p").findNext("p").a.text
        gross = movies.find("p", class_="sort-num_votes-visible").find_all("span")[-1].get_text()
        sheet.append([rank, movie_name, rating, year, story, gross])
except Exception as e:
    print(e)

wb.save(r"C:\Users\rknav\Downloads\Movie_list.xlsx")
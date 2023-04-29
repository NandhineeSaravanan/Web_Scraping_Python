from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Movie List"
sheet.append(['S.No', 'Movie_Name', 'Genre',
             'Rating', 'Story', 'Director', 'Vote', 'Gross'])

try:
    Response = requests.get("https://www.imdb.com/search/title/?genres=horror&sort=user_rating,desc&title_type=feature&num_votes=25000,&pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=94365f40-17a1-4450-9ea8-01159990ef7f&pf_rd_r=E42JM87E4EZ8PET4X1MN&pf_rd_s=right-6&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_gnr_12")
    soup = BeautifulSoup(Response.text, 'html.parser')
    # print(soup)

    movies = soup.find("div", class_="lister-list").find_all("div",
                                                             class_="lister-item mode-advanced")

    for movie in movies:
        rank = movie.find("h3").span.text.replace(".", "")
        name = movie.find("h3").a.text
        genre = movie.find(
            "p", class_="text-muted").find_all("span")[-1].get_text()
        rate = movie.find("div", class_="ratings-bar").strong.text
        story = movie.find("p").findNext("p", class_="text-muted").get_text()
        director = movie.find("p").findNext("p").findNext("p").a.text
        vote = movie.find(
            "p", class_="sort-num_votes-visible").find_all("span")[1].get_text()
        gross = movie.find(
            "p", class_="sort-num_votes-visible").find_all("span")[-1].get_text()
        print(rank, name, genre, rate, story, director, vote, gross)
        sheet.append([rank, name, genre, rate, story, director, vote, gross])

except Exception as e:
    print(e)

excel.save("Scraped_Data.xlsx")

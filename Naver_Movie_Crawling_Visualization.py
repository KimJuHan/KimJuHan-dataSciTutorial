import urllib.request
from bs4 import BeautifulSoup
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

matplotlib.rc('font', family="NanumBarunGothic")

class Review:
    def __init__(self, comment, date, star, good, bad):
        self.comment = comment
        self.date = date
        self.star = star
        self.good = good
        self.bad = bad

    def show(self):
        print("내용: " + self.comment +
              "\n날짜" + self.date +
              "\n별점" + self.star +
              "\n좋아요" + self.good +
              "\n싫어요" + self.bad
              )

def crawl(url):
    soup = BeautifulSoup(urllib.request.urlopen(url).read(), "html.parser")
    review_list = []
    title = soup.find('h3', class_="h_movie").find('a').text
    div = soup.find('div', class_="score_result")
    data_list = div.select("ul > li")

    for review in data_list:
        star = review.find("div", class_="star_score").text.strip()
        reply = review.find("div", class_="score_reple")
        comment = reply.find("p").text.strip()
        date = reply.select("dt > em")[1].text.strip()
        good = review.find("a", class_="_sympathyButton").select("strong")[0].text.strip()
        bad = review.find("a", class_="_notSympathyButton").select("strong")[0].text.strip()
        review_list.append(Review(comment, date, star, good, bad))

    return title, review_list

def get_summary(review_list):
    star_list = []
    good_list = []
    bad_list = []

    for review in review_list:
        star_list.append(int(review.star))
        good_list.append(int(review.good))
        bad_list.append(int(review.bad))

    star_series = pd.Series(star_list)
    good_series = pd.Series(good_list)
    bad_series = pd.Series(bad_list)

    summary = pd.DataFrame({
        'Star' : star_series,
        'Good' : good_series,
        'Bad' : bad_series,
        'Score' : good_series/(good_series + bad_series)
    })

    return summary

def movie_compare(review_lists):
    count = 1
    x = []
    y = []
    for movie, review_list in review_lists:
        x.append(count)
        summary = get_summary(review_list)
        summary = summary[summary['Score'] > 0.8]
        y.append(summary['Star'].mean())
        count += 1

    plt.bar(x, y)
    plt.title('영화 별점 비교')
    plt.xlabel('영화 번호')
    plt.ylabel('신뢰성 별점 평균')
    plt.show()


movie_code_list = [136900, 167657, 174321, 184859, 167391]

review_lists = []
for i in movie_code_list:
    title, review_list = crawl("https://movie.naver.com/movie/bi/mi/basic.nhn?code=" + str(i))
    summary = get_summary(review_list)
    print("[ %s ]" % (title))
    print(summary)
    review_lists.append((title, review_list))

movie_compare(review_lists)
import requests
from bs4 import BeautifulSoup
import json
import openpyxl
from random import choices
import datetime
from collections import namedtuple

COURSE_COUNT = 20

def get_courses_list(count=20):
    resp = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    bs = BeautifulSoup(resp.text, features='lxml')
    for course_url in choices(bs.find_all('loc'), k=count):
        yield course_url.get_text()


def get_course_info(course_url):
    resp = requests.get(course_url)
    course_text_info = resp.text.encode('iso-8859-1').decode('utf8')
    bs = BeautifulSoup(course_text_info, features='html.parser')
    course_data = json.loads(
        bs.select_one('script[type="application/ld+json"]').get_text()
    )
    course_name = course_data['@graph'][2]['name']
    start_date = course_data['@graph'][2]['hasCourseInstance']['startDate']
    end_date = course_data['@graph'][2]['hasCourseInstance']['endDate']
    weeks_count = get_weeks_count(start_date, end_date)
    rating = course_data['@graph'][1]['aggregateRating']['ratingValue']
    language = course_data['@graph'][2]['inLanguage']
    course_info = namedtuple(
        'course_info',
        'name language start_date weeks_count rating'
    )
    return course_info(course_name, language, start_date, weeks_count, rating)


def get_weeks_count(start_date, end_date):
    start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d')
    end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d')
    return round((end_date-start_date).days/7, ndigits=1)


def output_courses_info_to_xlsx(filepath, courses_info):
    book = openpyxl.Workbook()
    courses_sheet = book.create_sheet(title='Данные с Coursera')
    book.active = courses_sheet
    courses_sheet.append([
        'Название',
        'Язык',
        'Дата начала',
        'Продолжительность (недель)',
        'Средняя оценка'
    ])
    for course_info in courses_info:
        courses_sheet.append(course_info)
    book.save(filepath)


def main():
    courses_info = []
    try:
        for course_index, course_url in enumerate(get_courses_list(), start=1):
            print('Просмотрен {} из {}'.format(
                course_index,
                COURSE_COUNT
            ))
            courses_info.append(get_course_info(course_url))
        output_courses_info_to_xlsx('courses_info.xlsx', courses_info)
    except requests.exceptions.ConnectionError:
        exit("Ошибка соединения")


if __name__ == '__main__':
    main()

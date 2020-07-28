import json
import requests
import xlwt
import re
import datetime
import os


class YiBan:
    def __init__(self):
        with open('./config.json', 'r') as f:
            config = json.load(f)
        self.headers = config['headers']
        self.form_data = config['form_data']
        self.data_need = config['data']
        self.url = 'http://www.yiban.cn/forum/article/listAjax'
        self.date_now = str(datetime.datetime.now())
        self.date_now = self.date_now.split(' ')[0]

    def get_topics_by_num(self, num=10, sava_path='./yiban.xls'):
        self.form_data['size'] = num
        self.data = self._get_web()
        self.item_list = self._parse_data(self.data)
        self._save_data(self.item_list, save_path=sava_path)

    def get_topics_by_date(self, date, step=50, sava_path='./yiban.xls'):
        """
        get_topics_by_date
        assume 50 topics (step) a month
        :param date: format like: 20200728, 2020 07 28, 2020/07/28, 2020.07.28 2020-07-28
        :param step: number of topics posted a month
        :param sava_path:
        :return:
        """
        # 07-27 17:38 -> 2020-07-27
        if len(date) == 8:
            date = [date[0:4], date[4:6], date[6:8]]
            date = '-'.join(date)
        date = re.sub('[\./ ]', '-', date)
        check = re.findall('(\d{4}-\d{2}-\d{2})', date)
        if len(check) == 0:
            raise Exception('invaid input!')

        get = False
        batch = step
        item_list = list()
        while get is not True:
            self.form_data['size'] = batch
            web_data = self._get_web()
            item_list = self._parse_data(web_data)
            last_item = item_list[-1]
            last_item_create_time = last_item['createTime']
            last_item_create_date = last_item_create_time.split(' ')[0]
            if len(last_item_create_date) == 5:
                last_item_create_date = self.date_now[0:5] + last_item_create_date
            if len(last_item_create_date) != 10:
                raise Exception('date format error!')
            if self._date_compare(date, last_item_create_date):
                get = True
            else:
                batch += step
        item_list = self._contract_date(date, item_list)
        self._save_data(item_list, sava_path)

    def _contract_date(self, date, item_list):
        pos = 0
        found = False
        for item in item_list:
            item_create_time = item['createTime']
            item_create_date = item_create_time.split(' ')[0]
            if len(item_create_date) == 5:
                item_create_date = self.date_now[0:5] + item_create_date
            if len(item_create_date) != 10:
                raise Exception('date format error!')

            if self._date_compare(date, item_create_date):
                found = True
                break
            pos += 1

        if found:
            item_list = item_list[0:pos]
            return item_list
        else:
            raise Exception('No Found!')

    def _date_compare(self, date1, date2, fmt='%Y-%m-%d') -> bool:
        """
        比较两个真实日期之间的大小，date1 > date2 则返回True
        :param date1:
        :param date2:
        :param fmt:
        :return:
        """
        zero = datetime.datetime.fromtimestamp(0)
        try:
            d1 = datetime.datetime.strptime(str(date1), fmt)
        except:
            d1 = zero
        try:
            d2 = datetime.datetime.strptime(str(date2), fmt)
        except:
            d2 = zero
        return d1 > d2

    def _get_web(self):
        data = requests.post(url=self.url, data=self.form_data, headers=self.headers)
        data.encoding = 'utf-8'
        return data.json()

    def _parse_data(self, data):
        self.need_data = list()
        for attr, need in self.data_need.items():
            if need:
                self.need_data.append(attr)

        items = data['data']['list']
        item_list = list()
        for item in items:
            single_data = dict()
            for attr in self.need_data:
                single_data[attr] = item[attr]
            item_list.append(single_data)

        return item_list

    def _save_data(self, data, save_path):
        xls = xlwt.Workbook()
        table = xls.add_sheet('Sheet1')
        for i in range(len(self.need_data)):
            table.write(0, i, self.need_data[i])

        row = 1
        for item in data:
            col = 0
            for attr in self.need_data:
                table.write(row, col, item[attr])
                col += 1
            row += 1

        xls.save(save_path)
        print(f'Save {len(data)} posts successfully!')
        if 'images' in self.need_data:
            print('downloading images...')
            self._save_images(data)
            print('downloading successfully!')


    def _save_images(self, data):
        if not os.path.exists('./images'):
            os.mkdir('./images')
        for item in data:
            title = item['title']
            title = re.sub('[\\/:\*?"<>|]', '_', title)
            if not os.path.exists('./images/'+title):
                os.mkdir('./images/'+title)

            images = item['images']
            if len(images) == 0:
                continue
            for image in images:
                index = 1
                img = requests.get(image)
                img_name = image.split('/')[-1]
                with open('./images/'+title+'/'+str(img_name), 'wb') as f:
                    f.write(img.content)
                index += 1


if __name__ == '__main__':
    yiban = YiBan()
    # yiban.get_topics_by_num()
    yiban.get_topics_by_date('2020-07-01')


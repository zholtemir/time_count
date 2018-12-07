
# coding: utf-8

# In[655]:


import pandas as pd


# In[656]:


df = pd.read_excel('Для 4го отчета (датасет).xlsx')


# In[703]:


df.head()


# In[658]:


df['DT_DECISION']= pd.to_datetime(df['DT_DECISION'], dayfirst= True)
df['DT_LETTER']= pd.to_datetime(df['DT_LETTER'], dayfirst= True)

df = df.sort_values(by='DT_LETTER')


# # Список выходных дней

# In[659]:


auto_weekends = pd.read_excel('Выходные для АВТО.xlsx')

rest_weekends = pd.read_excel('Выходные для НЕ АВТО.xlsx')


# # Разбивка на недели

# In[660]:


from datetime import timedelta
from datetime import datetime

def week_range(date):
    """Find the first/last day of the week for the given day.
    Assuming weeks start on Sunday and end on Saturday.

    Returns a tuple of ``(start_date, end_date)``.

    """
    # isocalendar calculates the year, week of the year, and day of the week.
    # dow is Mon = 1, Sat = 6, Sun = 7
    year, week, dow = date.isocalendar()

    # Find the first day of the week.
    if dow == 1:
        # Since we want to start with Sunday, let's test for that condition.
        start_date = date
    else:
        # Otherwise, subtract `dow` number days to get the first day
        start_date = date - timedelta(dow - 1)

    # Now, add 6 for the last day of the week (i.e., count up to Saturday)
    end_date = start_date + timedelta(6)

    #return (start_date.date(), end_date.date())
    string = ''
    
    if start_date.date().day < 10:
        string = string + '0' + str(start_date.date().day)
    else:
        string = string + str(start_date.date().day)
    
    if start_date.date().month < 10:
        string = string + '.' + '0' + str(start_date.date().month)
    else:
        string = string + '.' + str(start_date.date().month)
    
    string = string + '.' + str(start_date.date().year) + '-'
    
    if end_date.date().day < 10:
        string = string + '0' + str(end_date.date().day)
    else:
        string = string + str(end_date.date().day)
    
    if end_date.date().month < 10:
        string = string + '.' + '0' + str(end_date.date().month)
    else:
        string = string + '.' + str(end_date.date().month)    
    string = string + '.' + str(start_date.date().year)
    
    #return ( str(start_date.date().day) + '.' + str(start_date.date().month) + '.' + str(start_date.date().year) + 
    #       '-' + str(end_date.date().day) + '.' + str(end_date.date().month) + '.' + str(end_date.date().year))
    return string    


df['week'] = df['DT_LETTER'].apply(week_range)


# # Список проверок и источников

# In[661]:


check_types = {
1: 'АВТО',
2: 'Перевозчик грузов',
3: 'Парк ТС',
4: 'Имущество',
5: 'Кандидат',
6: 'Агент',
7: 'Брокер',
8: 'СТОА',
9: 'Контрагент',
10: 'Суброгация',
11: 'Ипотека',
12: 'Запросы СБ',
13: 'Запросы УРМ',
14: 'ДЗ',
15: 'Осмотрщик',
16: 'Дилер',
17: 'ЛПУ',
18: 'Тоталь',
19: 'ДМС',
20: 'Арбитражный управляющий',
21: 'НС',
22: 'Страхование ответственности',
23: 'Тендер'}

sources = {
1: 'DI',
2: 'Агентский',
3: 'Брокерский',
4: 'Дилерский',
5: 'Корпоративный',
6: 'Банковский',
7: 'Лизинг',
8: 'Интач',
9: 'ЦОК',
10: 'Агентский',
11: 'Брокерский',
12: 'Дилерский',
13: 'Корпоративный',
14: 'Банковский',
15: 'Лизинг',
16: 'Тендер',
17: 'Штатный кандидат',
18: 'Аутсорсинг',
19: 'Стажер',
20: 'Тинькофф',
21: 'Иное',
0: 'Иное'}

df['Check_type'] = df['CHECK_TYPE_ID'].map(check_types)

df['Source'] = df['SRC_ID'].map(sources)


# # Класс BusinessHours

# In[662]:


import datetime


class BusinessHours:

    def __init__(self, datetime1, datetime2, worktiming=[8, 20]
                , weekends=[6, 7]):
        self.weekends = weekends
        self.worktiming = worktiming
        self.datetime1 = datetime1
        self.datetime2 = datetime2
        self.day_hours = (self.worktiming[1]-self.worktiming[0])
        self.day_minutes = self.day_hours * 60 # minutes in a work day

    def getdays(self):
        return int(self.getminutes() / self.day_minutes)
    
    def gethours(self):
        return int(self.getminutes() / 60)

    def getminutes(self):
        """
        Return the difference in minutes.
        """
        # 
        dt_start = self.datetime1  # datetime of start
        dt_end = self.datetime2    # datetime of end
        worktime_in_seconds = 0

        if dt_start.date() == dt_end.date():
            # starts and ends on same workday
            full_days = 0
            if self.is_weekend(dt_start):
                return 0
            else:
                if dt_start.hour < self.worktiming[0]:
                    # 
                    dt_start = datetime.datetime(
                        year=dt_start.year,
                        month=dt_start.month,
                        day=dt_start.day,
                        hour=self.worktiming[0],
                        minute=0)
                if dt_start.hour >= self.worktiming[1] or                         dt_end.hour < self.worktiming[0]:
                    return 0
                if dt_end.hour >= self.worktiming[1]:
                    dt_end = datetime.datetime(
                        year=dt_end.year,
                        month=dt_end.month,
                        day=dt_end.day,
                        hour=self.worktiming[1],
                        minute=0)
                worktime_in_seconds = (dt_end-dt_start).total_seconds()
        elif (dt_end-dt_start).days < 0:
            # ends before start
            return 0
        else:
            # start and ends on different days
            current_day = dt_start  # 
            while not current_day.date() == dt_end.date():
                if not self.is_weekend(current_day):
                    if current_day == dt_start:
                        # 
                        if current_day.hour < self.worktiming[0]:
                            # starts before the work day
                            worktime_in_seconds += self.day_minutes*60  # add 1 full work day
                        elif current_day.hour >= self.worktiming[1]:
                            pass  # 
                        else:
                            # 
                            dt_currentday_close = datetime.datetime(
                                year=dt_start.year,
                                month=dt_start.month,
                                day=dt_start.day,
                                hour=self.worktiming[1],
                                minute=0)
                            worktime_in_seconds += (dt_currentday_close
                                         - dt_start).total_seconds()
                    else:
                        # 
                        worktime_in_seconds += self.day_minutes*60
                current_day += datetime.timedelta(days=1)  # next day
            # Time on the last day
            if not self.is_weekend(dt_end):
                if dt_end.hour >= self.worktiming[1]:  # finish after close
                    # Add a full day
                    worktime_in_seconds += self.day_minutes*60
                elif dt_end.hour < self.worktiming[0]:  # close before opening
                    pass  # no time added
                else:
                    #
                    dt_end_open = datetime.datetime(
                        year=dt_end.year,
                        month=dt_end.month,
                        day=dt_end.day,
                        hour=self.worktiming[0],
                        minute=0)
                    worktime_in_seconds += (dt_end-dt_end_open).total_seconds()
        return int(worktime_in_seconds/60)

    def is_weekend(self, datetime):
        """
        Returns True if datetime is a weekend.
        """
        #for weekend in self.weekends:
        if datetime.date() in auto_weekends['DATE'].dt.date.values:
            #if datetime.isoweekday() == weekend:
            return True
        return False


# # Класс BusinessHours для дургих типов проверок

# In[663]:


import datetime

class RestBusinessHours:

    def __init__(self, datetime1, datetime2, worktiming=[8, 20]
                , weekends=[6, 7]):
        self.weekends = weekends
        self.worktiming = worktiming
        self.datetime1 = datetime1
        self.datetime2 = datetime2
        self.day_hours = (self.worktiming[1]-self.worktiming[0])
        self.day_minutes = self.day_hours * 60 # minutes in a work day

    def getdays(self):
        return int(self.getminutes() / self.day_minutes)
    
    def gethours(self):
        return int(self.getminutes() / 60)

    def getminutes(self):
        """
        Return the difference in minutes.
        """
        # 
        dt_start = self.datetime1  # datetime of start
        dt_end = self.datetime2    # datetime of end
        worktime_in_seconds = 0

        if dt_start.date() == dt_end.date():
            # starts and ends on same workday
            full_days = 0
            if self.is_weekend(dt_start):
                return 0
            else:
                if dt_start.hour < self.worktiming[0]:
                    # set start time to opening hour
                    dt_start = datetime.datetime(
                        year=dt_start.year,
                        month=dt_start.month,
                        day=dt_start.day,
                        hour=self.worktiming[0],
                        minute=0)
                if dt_start.hour >= self.worktiming[1] or                         dt_end.hour < self.worktiming[0]:
                    return 0
                if dt_end.hour >= self.worktiming[1]:
                    dt_end = datetime.datetime(
                        year=dt_end.year,
                        month=dt_end.month,
                        day=dt_end.day,
                        hour=self.worktiming[1],
                        minute=0)
                worktime_in_seconds = (dt_end-dt_start).total_seconds()
        elif (dt_end-dt_start).days < 0:
            # ends before start
            return 0
        else:
            # start and ends on different days
            current_day = dt_start  # 
            while not current_day.date() == dt_end.date():
                if not self.is_weekend(current_day):
                    if current_day == dt_start:
                        # 
                        if current_day.hour < self.worktiming[0]:
                            # starts before the work day
                            worktime_in_seconds += self.day_minutes*60  # 
                        elif current_day.hour >= self.worktiming[1]:
                            pass  # 
                        else:
                            # 
                            dt_currentday_close = datetime.datetime(
                                year=dt_start.year,
                                month=dt_start.month,
                                day=dt_start.day,
                                hour=self.worktiming[1],
                                minute=0)
                            worktime_in_seconds += (dt_currentday_close
                                         - dt_start).total_seconds()
                    else:
                        # 
                        worktime_in_seconds += self.day_minutes*60
                current_day += datetime.timedelta(days=1)  # next day
            # Time on the last day
            if not self.is_weekend(dt_end):
                if dt_end.hour >= self.worktiming[1]:  # finish after close
                    # Add a full day
                    worktime_in_seconds += self.day_minutes*60
                elif dt_end.hour < self.worktiming[0]:  # close before opening
                    pass  # no time added
                else:
                    # Add time since opening
                    dt_end_open = datetime.datetime(
                        year=dt_end.year,
                        month=dt_end.month,
                        day=dt_end.day,
                        hour=self.worktiming[0],
                        minute=0)
                    worktime_in_seconds += (dt_end-dt_end_open).total_seconds()
        return int(worktime_in_seconds/60)

    def is_weekend(self, datetime):
        """
        Returns True if datetime is a weekend.
        """
        #for weekend in self.weekends:
        if datetime.date() in rest_weekends['DATE'].dt.date.values:
            #if datetime.isoweekday() == weekend:
            return True
        return False


# # Считаем время обработки заявки

# In[664]:


get_ipython().run_cell_magic('time', '', "\ndiff = []\nfor i, j, k in zip(df['DT_LETTER'], df['DT_DECISION'], df['CHECK_TYPE_ID']):\n    try:\n        if k == 1:\n            diff.append(BusinessHours(i, j).getminutes()/60)\n        else:\n            diff.append(RestBusinessHours(i, j).getminutes()/60)\n    except :\n        diff.append('NAN')\n        \ndf['time_diff'] = diff")


# In[665]:


df = df[df['time_diff'] != 'NAN']

df['time_diff'] = pd.to_numeric(df['time_diff'])


# In[666]:


df.shape


# In[667]:


len(diff)


# In[668]:


df.head()


# In[669]:


def cut_date(x):
    return pd.to_datetime(x[:10], dayfirst= True)

def week_stop(x):
    return pd.to_datetime(x[11:], dayfirst= True)


# # Все АВТО

# In[670]:


AUTO = df[(df['Check_type'] == 'АВТО')]

#AUTO =  df[(df['Check_type'] == 'АВТО') &  (df['Source'] != 'DI')]

AUTO['is_long'] = np.where(AUTO['time_diff'] > 3, 1, 0) # больше ли регламентированного времени

AUTO_grouped = AUTO.groupby('week')

mean_median_AUTO = AUTO_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_AUTO = (AUTO_grouped['is_long'].sum() / AUTO_grouped['is_long'].count()).reset_index()

AUTO_final = pd.merge(mean_median_AUTO, max_portion_AUTO, how ='inner', on='week') # concat two tables

AUTO_final['week_start'] = AUTO_final['week'].apply(cut_date) # apply function cut_date to sort values

AUTO_final['week_stop'] = AUTO_final['week'].apply(week_stop)
 
AUTO_final = AUTO_final.sort_values(by='week_start')


# In[671]:


AUTO_final.head()


# # Авто агентские

# In[672]:


AGENT = df[(df['Check_type'] == 'АВТО') & (df['Source'] == 'Агентский')  & (df['CATEGORY_ID'] == 'ФЛ')]

AGENT['is_long'] = np.where(AGENT['time_diff'] > 1, 1, 0) # больше ли регламентированного времени

AGENT_grouped = AGENT.groupby('week')

mean_median_AGENT = AGENT_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_AGENT = (AGENT_grouped['is_long'].sum() / AGENT_grouped['is_long'].count()).reset_index()

AGENT_final = pd.merge(mean_median_AGENT, max_portion_AGENT, how ='inner', on='week') # concat two tables

AGENT_final['week_start'] = AGENT_final['week'].apply(cut_date) # apply function cut_date to sort values

AGENT_final['week_stop'] = AGENT_final['week'].apply(week_stop)
 
AGENT_final = AGENT_final.sort_values(by='week_start')


# In[673]:


AGENT_final.head()


# # DI

# In[674]:


DI = df[(df['Check_type'] == 'АВТО') & (df['Source'] == 'DI') & (df['CATEGORY_ID'] == 'ФЛ')]

DI['is_long'] = np.where(DI['time_diff'] > 1, 1, 0) # больше ли регламентированного времени

DI_grouped = DI.groupby('week')

mean_median_DI = DI_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_DI = (DI_grouped['is_long'].sum() / DI_grouped['is_long'].count()).reset_index()

DI_final = pd.merge(mean_median_DI, max_portion_DI, how ='inner', on='week') # concat two tables

DI_final['week_start'] = DI_final['week'].apply(cut_date) # apply function cut_date to sort values

DI_final['week_stop'] = DI_final['week'].apply(week_stop)
 
DI_final = DI_final.sort_values(by='week_start')


# In[675]:


DI_final.head()


# # Дилеры

# In[676]:


DEALER = df[(df['Check_type'] == 'АВТО') & (df['Source'] == 'Дилерский')  & (df['CATEGORY_ID'] == 'ФЛ')]

DEALER['is_long'] = np.where(DEALER['time_diff'] > 3, 1, 0) # больше ли регламентированного времени

DEALER_grouped = DEALER.groupby('week')

mean_median_DEALER = DEALER_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_DEALER = (DEALER_grouped['is_long'].sum() / DEALER_grouped['is_long'].count()).reset_index()

DEALER_final = pd.merge(mean_median_DEALER, max_portion_DEALER, how ='inner', on='week') # concat two tables

DEALER_final['week_start'] = DEALER_final['week'].apply(cut_date) # apply function cut_date to sort values

DEALER_final['week_stop'] = DEALER_final['week'].apply(week_stop)
 
DEALER_final = DEALER_final.sort_values(by='week_start')


# In[677]:


DEALER_final.head()


# # БРОКЕРЫ

# In[678]:


BROKER = df[(df['Check_type'] == 'АВТО') & (df['Source'] == 'Брокерский')  & (df['CATEGORY_ID'] == 'ФЛ')]

BROKER['is_long'] = np.where(BROKER['time_diff'] > 3, 1, 0) # больше ли регламентированного времени

BROKER_grouped = BROKER.groupby('week')

mean_median_BROKER = BROKER_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_BROKER = (BROKER_grouped['is_long'].sum() / BROKER_grouped['is_long'].count()).reset_index()

BROKER_final = pd.merge(mean_median_BROKER, max_portion_BROKER, how ='inner', on='week') # concat two tables

BROKER_final['week_start'] = BROKER_final['week'].apply(cut_date) # apply function cut_date to sort values

BROKER_final['week_stop'] = BROKER_final['week'].apply(week_stop)
 
BROKER_final = BROKER_final.sort_values(by='week_start')


# In[679]:


BROKER_final.head()


# # Парк ТС

# In[680]:


PARK_TS = df[(df['Check_type'] == 'Парк ТС')]

PARK_TS['is_long'] = np.where(PARK_TS['time_diff'] > 24, 1, 0) # больше ли регламентированного времени

PARK_TS_grouped = PARK_TS.groupby('week')

mean_median_PARK_TS = PARK_TS_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_PARK_TS = (PARK_TS_grouped['is_long'].sum() / PARK_TS_grouped['is_long'].count()).reset_index()

PARK_TS_final = pd.merge(mean_median_PARK_TS, max_portion_PARK_TS, how ='inner', on='week') # concat two tables

PARK_TS_final['week_start'] = PARK_TS_final['week'].apply(cut_date) # apply function cut_date to sort values

PARK_TS_final['week_stop'] = PARK_TS_final['week'].apply(week_stop)
 
PARK_TS_final = PARK_TS_final.sort_values(by='week_start')


# In[681]:


PARK_TS_final.head()


# # Перевозчики грузов

# In[682]:


GRUZ = df[(df['Check_type'] == 'Перевозчик грузов')]

GRUZ['is_long'] = np.where(GRUZ['time_diff'] > 4, 1, 0) # больше ли регламентированного времени

GRUZ_grouped = GRUZ.groupby('week')

mean_median_GRUZ = GRUZ_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_GRUZ = (GRUZ_grouped['is_long'].sum() / GRUZ_grouped['is_long'].count()).reset_index()

GRUZ_final = pd.merge(mean_median_GRUZ, max_portion_GRUZ, how ='inner', on='week') # concat two tables

GRUZ_final['week_start'] = GRUZ_final['week'].apply(cut_date) # apply function cut_date to sort values

GRUZ_final['week_stop'] = GRUZ_final['week'].apply(week_stop)
 
GRUZ_final = GRUZ_final.sort_values(by='week_start')


# In[683]:


GRUZ_final.head()


# # Имущество ФЛ

# In[684]:


IMUSHESTVO_FL = df[(df['Check_type'] == 'Имущество') & (df['CATEGORY_ID'] == 'ФЛ')]

IMUSHESTVO_FL['is_long'] = np.where(IMUSHESTVO_FL['time_diff'] > 24, 1, 0) # больше ли регламентированного времени

IMUSHESTVO_FL_grouped = IMUSHESTVO_FL.groupby('week')

mean_median_IMUSHESTVO_FL = IMUSHESTVO_FL_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_IMUSHESTVO_FL = (IMUSHESTVO_FL_grouped['is_long'].sum() / IMUSHESTVO_FL_grouped['is_long'].count()).reset_index()

IMUSHESTVO_FL_final = pd.merge(mean_median_IMUSHESTVO_FL, max_portion_IMUSHESTVO_FL, how ='inner', on='week') # concat two tables

IMUSHESTVO_FL_final['week_start'] = IMUSHESTVO_FL_final['week'].apply(cut_date) # apply function cut_date to sort values

IMUSHESTVO_FL_final['week_stop'] = IMUSHESTVO_FL_final['week'].apply(week_stop)
 
IMUSHESTVO_FL_final = IMUSHESTVO_FL_final.sort_values(by='week_start')


# In[685]:


IMUSHESTVO_FL_final.head()


# # ИПОТЕКА

# In[686]:


IPOTEKA = df[(df['Check_type'] == 'Ипотека')]

IPOTEKA['is_long'] = np.where(IPOTEKA['time_diff'] > 8, 1, 0) # больше ли регламентированного времени

IPOTEKA_grouped = IPOTEKA.groupby('week')

mean_median_IPOTEKA = IPOTEKA_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_IPOTEKA = (IPOTEKA_grouped['is_long'].sum() / IPOTEKA_grouped['is_long'].count()).reset_index()

IPOTEKA_final = pd.merge(mean_median_IPOTEKA, max_portion_IPOTEKA, how ='inner', on='week') # concat two tables

IPOTEKA_final['week_start'] = IPOTEKA_final['week'].apply(cut_date) # apply function cut_date to sort values

IPOTEKA_final['week_stop'] = IPOTEKA_final['week'].apply(week_stop)
 
IPOTEKA_final = IPOTEKA_final.sort_values(by='week_start')


# In[687]:


IPOTEKA_final.head()


# # СТРАХОВАТЕЛИ ЮЛ

# In[688]:


imushestvo_not_fl = df[(df['Check_type'] == 'Имущество') & (df['CATEGORY_ID'] != 'ФЛ')]

rest = df[(df['Check_type'] == 'Страхование ответственности') | (df['Check_type'] == 'Арбитражный управляющий')]

STRAHOVATELI = pd.concat([imushestvo_not_fl, rest])

STRAHOVATELI['is_long'] = np.where(STRAHOVATELI['time_diff'] > 24, 1, 0) # больше ли регламентированного времени

STRAHOVATELI_grouped = STRAHOVATELI.groupby('week')

mean_median_STRAHOVATELI = STRAHOVATELI_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_STRAHOVATELI = (STRAHOVATELI_grouped['is_long'].sum() / STRAHOVATELI_grouped['is_long'].count()).reset_index()

STRAHOVATELI_final = pd.merge(mean_median_STRAHOVATELI, max_portion_STRAHOVATELI, how ='inner', on='week') # concat two tables

STRAHOVATELI_final['week_start'] = STRAHOVATELI_final['week'].apply(cut_date) # apply function cut_date to sort values

STRAHOVATELI_final['week_stop'] = STRAHOVATELI_final['week'].apply(week_stop)
 
STRAHOVATELI_final = STRAHOVATELI_final.sort_values(by='week_start')


# In[689]:


STRAHOVATELI_final.head()


# # Кандидат

# In[690]:


candidate = df[(df['Check_type'] == 'Кандидат')]

osmotrshik = df[(df['Check_type'] == 'Осмотрщик') & (df['CATEGORY_ID'] == 'ФЛ')]

CANDIDATE = pd.concat([candidate, osmotrshik])

CANDIDATE['is_long'] = np.where(CANDIDATE['time_diff'] > 24, 1, 0) # больше ли регламентированного времени

CANDIDATE_grouped = CANDIDATE.groupby('week')

mean_median_CANDIDATE = CANDIDATE_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_CANDIDATE = (CANDIDATE_grouped['is_long'].sum() / CANDIDATE_grouped['is_long'].count()).reset_index()

CANDIDATE_final = pd.merge(mean_median_CANDIDATE, max_portion_CANDIDATE, how ='inner', on='week') # concat two tables

CANDIDATE_final['week_start'] = CANDIDATE_final['week'].apply(cut_date) # apply function cut_date to sort values

CANDIDATE_final['week_stop'] = CANDIDATE_final['week'].apply(week_stop)
 
CANDIDATE_final = CANDIDATE_final.sort_values(by='week_start')


# In[691]:


CANDIDATE_final.head()


# # Агент ФЛ и ИП

# In[692]:


AGENT_FL = df[(df['Check_type'] == 'Агент') & (df['CATEGORY_ID'] != 'ЮЛ') ]

AGENT_FL['is_long'] = np.where(AGENT_FL['time_diff'] > 24, 1, 0) # больше ли регламентированного времени

AGENT_FL_grouped = AGENT_FL.groupby('week')

mean_median_AGENT_FL = AGENT_FL_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_AGENT_FL = (AGENT_FL_grouped['is_long'].sum() / AGENT_FL_grouped['is_long'].count()).reset_index()

AGENT_FL_final = pd.merge(mean_median_AGENT_FL, max_portion_AGENT_FL, how ='inner', on='week') # concat two tables

AGENT_FL_final['week_start'] = AGENT_FL_final['week'].apply(cut_date) # apply function cut_date to sort values

AGENT_FL_final['week_stop'] = AGENT_FL_final['week'].apply(week_stop)
 
AGENT_FL_final = AGENT_FL_final.sort_values(by='week_start')


# In[693]:


AGENT_FL_final.head()


# # Посредники

# In[694]:


part_1 = df[(df['Check_type'] == 'Контрагент') | (df['Check_type'] == 'Брокер')| (df['Check_type'] == 'СТОА') |
  (df['Check_type'] == 'Дилер') | (df['Check_type'] == 'ДМС') | (df['Check_type'] == 'ЛПУ')
  |(df['Check_type'] == 'Тендер') | (df['Check_type'] == 'Тоталь')]

part_2 = df[(df['Check_type'] == 'Осмотрщик') & (df['CATEGORY_ID'] != 'ФЛ')]

part_3 = df[(df['Check_type'] == 'Агент') & (df['CATEGORY_ID'] == 'ЮЛ')] 

POSREDNIKI = pd.concat([part_1, part_2, part_3])

POSREDNIKI['is_long'] = np.where(POSREDNIKI['time_diff'] > 40, 1, 0) # больше ли регламентированного времени

POSREDNIKI_grouped = POSREDNIKI.groupby('week')

mean_median_POSREDNIKI = POSREDNIKI_grouped['time_diff'].agg(['mean', 'median', 'min', 'max']).reset_index()

max_portion_POSREDNIKI = (POSREDNIKI_grouped['is_long'].sum() / POSREDNIKI_grouped['is_long'].count()).reset_index()

POSREDNIKI_final = pd.merge(mean_median_POSREDNIKI, max_portion_POSREDNIKI, how ='inner', on='week') # concat two tables

POSREDNIKI_final['week_start'] = POSREDNIKI_final['week'].apply(cut_date) # apply function cut_date to sort values

POSREDNIKI_final['week_stop'] = POSREDNIKI_final['week'].apply(week_stop)
 
POSREDNIKI_final = POSREDNIKI_final.sort_values(by='week_start')


# In[695]:


POSREDNIKI_final.head()


# In[701]:


get_ipython().run_cell_magic('time', '', 'writer = pd.ExcelWriter("4 otchet KKK.xlsx")\n\nAUTO_final.to_excel(writer,\'АВТО ФЛ\', index=False)\n\nAGENT_final.to_excel(writer,\'Авто Агентские\', index=False)\n\nDI_final.to_excel(writer,\'Авто DI\', index=False)\n\nDEALER_final.to_excel(writer,\'Авто Дилерский\', index=False)\n\nBROKER_final.to_excel(writer,\'Авто Брокеры\', index=False)\n\nPARK_TS_final.to_excel(writer,\'Парк ТС\', index=False)\n\nGRUZ_final.to_excel(writer,\'Перевозчики грузов\', index=False)\n\nIMUSHESTVO_FL_final.to_excel(writer,\'Имущество ФЛ\', index=False)\n\nIPOTEKA_final.to_excel(writer,\'Ипотека\', index=False)\n\nSTRAHOVATELI_final.to_excel(writer,\'Страхователи ЮЛ\', index=False)\n\nCANDIDATE_final.to_excel(writer,\'Кандидат\', index=False)\n\nAGENT_FL_final.to_excel(writer,\'Агенты ФЛ и ИП\', index=False)\n\nPOSREDNIKI_final.to_excel(writer,\'Посредники\', index=False)\n\ndf.to_excel(writer,\'Все записи\', index=False)\n\nwriter.save()')


# In[697]:


#df.shape


# In[698]:


#df[(df['Check_type'] == 'АВТО') & (df['time_diff'] > 8)]


# In[699]:


#import cx_Oracle


# In[700]:


"""
URM = 
  (DESCRIPTION = 
    (ADDRESS = 
      (PROTOCOL = TCP)
      (HOST = renmskorurm02.mos.renins.com)
      (PORT = 1521)
    )
    (CONNECT_DATA = 
      (SERVICE_NAME = URM)
    )
  )
"""


# In[702]:


"""
host = 'renmskorurm02.mos.renins.com'
port = 1521
service_name = 'URM'

dsn_tns = cx_Oracle.makedsn(host, port, service_name)

db = cx_Oracle.connect('balbasa', 'Rjhbfylh', dsn_tns)
"""


# In[ ]:


"""
cx_Oracle.connect(user='muk12345', password='12345678q', dsn=None, mode=None, handle=None, pool=None, threaded=False, 
                  events=False, cclass=None, purity=None, newpassword=None, encoding=None, nencoding=None, edition=None,
                  appcontext=[], tag=None, matchanytag=None, shardingkey=[], supershardingkey=[])
"""


# In[ ]:


print (dsn_tns)


# In[704]:


get_ipython().system('pip install wheel')


# In[705]:


get_ipython().system('pip install pyinstaller')


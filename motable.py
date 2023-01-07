import pandas as pd
import datetime as dt

class MOTable:
    def __init__(self, file, age) -> None:
        ''' Initialisation of data frame without empty blocks and duplicates '''
        self.data_frame = pd.read_excel(file)
        self.data_frame = self.data_frame.dropna(axis=0)
        self.data_frame = self.data_frame.drop_duplicates(subset=['emiasid'])
        self.age = age
        self.current_year = dt.datetime.now()
        self.point_year = self.current_year - pd.Timedelta(days=age*365)


    def get_file_total_by_age(self):
        ''' Get the total number of people who are over or under specified years old separately for each MO '''
        self.data_frame['younger_than_' +
                        str(self.age)] = self.data_frame['birth_date'] > self.point_year
        self.data_frame['older_than_' +
                        str(self.age)] = self.data_frame['birth_date'] < self.point_year
        pivot_table = self.data_frame.pivot_table(
            columns='МО',
            values=[('younger_than_' + str(self.age)),
                    ('older_than_' + str(self.age))],
            aggfunc='sum')
        pivot_table.to_excel(
            'total_by_age.xlsx',
            sheet_name='Count by ' + str(self.age) + 'yr old in MO')

    def get_files_by_mo(self):
        self.data_frame = self.data_frame[self.data_frame[(
            'younger_than_' + str(self.age))] == True]
        self.data_frame['fullname'] = self.data_frame.apply(
            lambda x:
            x['family'] + ' ' +
            x['name'] + ' ' +
            x['patronimic'],
            axis=1)
        self.data_frame.insert(2, 'fullname', self.data_frame.pop('fullname'))
        self.data_frame = self.data_frame.drop(
            ['family',
             'name',
             'patronimic',
             'older_than_' + str(self.age),
             'younger_than_' + str(self.age)],
            axis=1)
        for i in self.data_frame['МО'].unique():
            splited_data_frame_by_mo = self.data_frame[self.data_frame['МО'] == i]
            splited_data_frame_by_mo.to_excel(i + '.xlsx', sheet_name='MO')


if __name__ == "__main__":
    motable = MOTable(file='File1.xlsx', age=20)
    motable.get_file_total_by_age()
    motable.get_files_by_mo()

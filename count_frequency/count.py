import pandas as pd


def main():
    # read excel file into pandas dataframe
    df = pd.read_excel('data.xlsx')

    # count number of occurrences for a column
    # returns a pandas series
    # https://stackoverflow.com/questions/20076195/what-is-the-most-efficient-way-of-counting-occurrences-in-pandas
    # https://towardsdatascience.com/getting-more-value-from-the-pandas-value-counts-aa17230907a6
    val_counts = df['ID'].value_counts()
    print(val_counts)

    # (optional) convert series to dataframe and name columns
    # https://stackoverflow.com/questions/28503445/assigning-column-names-to-a-pandas-series
    df1 = val_counts.to_frame('Frequency')
    df1.index.name = 'ID'
    print(df1)

    # output to excel
    # https://xlsxwriter.readthedocs.io/example_pandas_simple.html
    writer = pd.ExcelWriter('Frequency Count.xlsx', engine='xlsxwriter')
    df1.to_excel(writer)
    writer.save()


if __name__ == '__main__':
    main()

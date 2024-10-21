# %% [markdown]
# ### Step 1 - Import Archives and Librarys

# %%
import pandas as pd
import pathlib as pl
import win32com.client as win32
import pythoncom

# %%
emails = pd.read_excel(r'Databases/emails.xlsx')
stores = pd.read_csv(r'Databases/stores.csv', encoding='latin-1', sep=';')   # Using latin-1 encoding because the file has some characteres not recognized by utf-8
sales = pd.read_excel(r'Databases/sales.xlsx')

# %% [markdown]
# ### Step 2 - Define one sheet to each store and definne the indicator day
# 

# %%
# include store name to "sales"

sales = sales.merge(stores, on='Store ID')

# %%
store_dictionary = {}

for item in stores['Store']:
    store_dictionary[item] = sales.loc[sales['Store']==item, :]

# %%
# Find the latest day available and get its data

indicator_day = sales['Date'].max()

# %% [markdown]
# ### Step 3 - Save the sheets in the Backup folder
# 

# %%
# identify if folder already exists
backup_path = pl.Path(r'Backup')
backup_folder_files = backup_path.iterdir()
backup_name_list = []

for file in backup_folder_files:
    backup_name_list.append(file.name)

# %%
# save inside the respective folder
for store in store_dictionary:
    if store not in backup_name_list:
        new_folder = backup_path / store
        new_folder.mkdir()

    file_name = '{}_{}_{}.xlsx'.format(indicator_day.month, indicator_day.day, store)

    file_path = backup_path / store / file_name

    store_dictionary[store].to_excel(file_path)


# %% [markdown]
# ### Step 4 - Calculate the indicators (Revenue from the year, revenue from the last day in dataframe, product diversity from sales and average ticket from each store)
# 

# %%
# Define goals
revenue_year_goal = 1650000
revenue_day_goal = 1000
product_amount_year_goal = 120
product_amount_day_goal = 4
average_ticket_year_goal = 60000
average_ticket_day_goal = 500


# %%
for store in store_dictionary:
    store_sales_year = store_dictionary[store]
    store_sales_day = store_sales_year.loc[store_sales_year['Date']==indicator_day, :]

    # Revenue indicators
    revenue_year = store_sales_year['Final Value'].sum()
    revenue_day = store_sales_day['Final Value'].sum()

    # Product diversity indicators
    product_amount_year = len(store_sales_year['Product'].unique())
    product_amount_day = len(store_sales_day['Product'].unique())

    # Average ticket indicators
    sale_values_year = store_sales_year.groupby('Sale Code').sum('Final Value')
    sale_values_day = store_sales_day.groupby('Sale Code').sum('Final Value')
    average_ticket_year = sale_values_year['Final Value'].mean()
    average_ticket_day = sale_values_day['Final Value'].mean()

    # Send email to the store manager
    attachment = pl.Path.cwd() / backup_path / store / '{}_{}_{}.xlsx'.format(indicator_day.month, indicator_day.day, store)

    outlook = win32.Dispatch('outlook.application', pythoncom.CoInitialize())

    color_revenue_day = "green" if revenue_day >= revenue_day_goal else "red"
    color_revenue_year = "green" if revenue_year >= revenue_year_goal else "red"
    color_diversity_day = "green" if product_amount_day >= product_amount_day_goal else "red"
    color_diversity_year = "green" if product_amount_year >= product_amount_year_goal else "red"
    color_ticket_day = "green" if average_ticket_day >= average_ticket_day_goal else "red"
    color_ticket_year = "green" if average_ticket_year >= average_ticket_year_goal else "red"

    name = emails.loc[emails['Store']==store, 'Manager'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Store']==store, 'Email'].values[0]
    mail.Subject = 'Indicators - Date {}/{} ({})'.format(indicator_day.month, indicator_day.day, store)
    # Use mail.Body if you want to make a simple text mail.
    mail.HTMLBody = f'''
    <p>Hi {emails.loc[emails['Store']==store, 'Manager'].values[0]},</p>

    <p>Attached is the financial report for the <strong>{store}</strong> store from <strong>yesterday ({indicator_day.day}/{indicator_day.month})</strong>.</p>

    <p>&nbsp;</p>

    <p>Here is a brief summary:</p>

    <table align="center" border="1" cellpadding="1" cellspacing="1" style="width:500px">
        <thead>
            <tr>
                <th scope="col">Indicator</th>
                <th scope="col">Day Value</th>
                <th scope="col">Day Goal</th>
                <th scope="col">Day Scenario</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>
                <p>Revenue</p>
                </td>
                <td style="text-align: center;">R$ {revenue_day:.2f}</td>
                <td style="text-align: center;">R$ {revenue_day_goal:.2f}</td>
                <td style="text-align: center;"><font color="{color_revenue_day}">◙</td>
            </tr>
            <tr>
                <td>
                <p>Product Diversity</p>
                </td>
                <td style="text-align: center;">{product_amount_day}</td>
                <td style="text-align: center;">{product_amount_day_goal}</td>
                <td style="text-align: center;"><font color="{color_diversity_day}">◙</td>
            </tr>
            <tr>
                <td>
                <p>Average Ticket</p>
                </td>
                <td style="text-align: center;">R$ {average_ticket_day:.2f}</td>
                <td style="text-align: center;">R$ {average_ticket_day_goal:.2f}</td>
                <td style="text-align: center;"><font color="{color_ticket_day}">◙</td>
            </tr>
        </tbody>
    </table>
    <br>
    <table align="center" border="1" cellpadding="1" cellspacing="1" style="width:500px">
        <thead>
            <tr>
                <th scope="col">Indicator</th>
                <th scope="col">Year Value</th>
                <th scope="col">Year Goal</th>
                <th scope="col">Year Scenario</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>
                <p>Revenue</p>
                </td>
                <td style="text-align: center;">R$ {revenue_year:.2f}</td>
                <td style="text-align: center;">R$ {revenue_year_goal:.2f}</td>
                <td style="text-align: center;"><font color="{color_revenue_year}">◙</td>
            </tr>
            <tr>
                <td>
                <p>Product Diversity</p>
                </td>
                <td style="text-align: center;">{product_amount_year}</td>
                <td style="text-align: center;">{product_amount_year_goal}</td>
                <td style="text-align: center;"><font color="{color_diversity_year}">◙</td>
            </tr>
            <tr>
                <td>
                <p>Average Ticket</p>
                </td>
                <td style="text-align: center;">R$ {average_ticket_year:.2f}</td>
                <td style="text-align: center;">R$ {average_ticket_year_goal:.2f}</td>
                <td style="text-align: center;"><font color="{color_ticket_year}">◙</td>
            </tr>
        </tbody>
    </table>

    <p>&nbsp;</p>

    <hr />
    <p style="text-align: center;"><strong>Please note that this is an automated email. For any questions, please contact the headquarter.</strong></p>

    <p style="text-align: center;">&nbsp;</p>

    <p style="text-align: center;">Best regards,</p>

    <p style="text-align: center;">SIA - Store Indicators Automation.</p>
    '''

    attachment = pl.Path.cwd() / backup_path / store / '{}_{}_{}.xlsx'.format(indicator_day.month, indicator_day.day, store)
    mail.Attachments.Add(str(attachment))

    mail.Send()

# %% [markdown]
# ### Step 7 - Create ranking for the director of the company and saving to excel
# 

# %%
revenue_by_store = sales.groupby("Store").agg({
    "Final Value": "sum"
})
revenue_by_store = revenue_by_store.sort_values(by='Final Value', ascending=False)

sales_day = sales.loc[sales['Date']==indicator_day , :]
revenue_by_store_day = sales_day.groupby("Store")[['Store', 'Final Value']].agg({
    "Final Value": "sum"
})
revenue_by_store_day = revenue_by_store_day.sort_values(by='Final Value', ascending=False)


annual_ranking_file_name = '{}_{}_annual_ranking.xlsx'.format(indicator_day.month, indicator_day.day)
annual_ranking_file_path = backup_path / annual_ranking_file_name
revenue_by_store.to_excel(annual_ranking_file_path)

daily_ranking_file_name = '{}_{}_daily_ranking.xlsx'.format(indicator_day.month, indicator_day.day)
daily_ranking_file_path = backup_path / daily_ranking_file_name
revenue_by_store.to_excel(daily_ranking_file_path)


# %% [markdown]
# ### Step 8 - Send email to the director

# %%
outlook = win32.Dispatch('outlook.application', pythoncom.CoInitialize())

color_revenue_day = "green" if revenue_day >= revenue_day_goal else "red"
color_revenue_year = "green" if revenue_year >= revenue_year_goal else "red"
color_diversity_day = "green" if product_amount_day >= product_amount_day_goal else "red"
color_diversity_year = "green" if product_amount_year >= product_amount_year_goal else "red"
color_ticket_day = "green" if average_ticket_day >= average_ticket_day_goal else "red"
color_ticket_year = "green" if average_ticket_year >= average_ticket_year_goal else "red"

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Store']=='BOARD OF DIRECTORS', 'Email'].values[0]
mail.Subject = 'Ranking - Date {}/{}'.format(indicator_day.month, indicator_day.day)
mail.Body = f'''
Dear Sirs,

I hope this email finds you well.

Here it is the update on our store's performance. Yesterday, {revenue_by_store_day.index[0]} achieved the highest sales, generating a total of R$ {revenue_by_store_day.iloc[0, 0]:.2f}.

To provide you with a more comprehensive view of our performance, you will find attached two tables:

Annual Store Ranking: This table provides a ranking of all our stores based on their year-to-date sales performance.
Daily Store Ranking: This table presents a ranking of all our stores based on their sales performance for yesterday.

Please do not reply to this email, since it is automatically sent by our automation.

Best regards,

SIA - Store Indicators Automation.
'''

attachment = pl.Path.cwd() / annual_ranking_file_path
mail.Attachments.Add(str(attachment))
attachment = pl.Path.cwd() / daily_ranking_file_path
mail.Attachments.Add(str(attachment))

mail.Send()



# -*- coding: utf-8 -*-
"""
Updated on 11/16/2020

@author: Emma Mingyuan Zhu

- According to meeting with Kelly 8/18/2020, he provides some feedbacks in simplifying and expanding the wkly reports to all three banners.
    - Removing the middle and bottom charts of cohort & shoppers conversion analysis
    - Extending the first shoppers tracking charts to all banners (SR, PR, FG) and replcaing the shoppers metric to sales, since sales is the ultimate metric goal in perspective of our client
- Added Units as key Metrics per Chris' request, EZ 9/24/2020
- Added New & Both Sales/Units share in the final summary excel deck, EZ 11/03/2020
- Set docx table cell background and text color, EZ 11/16/2020

"""

def connectYB(database_name,uid,pwd):
    import platform
    path=platform.system()

    if path=='Windows':
        connectstring="DRIVER={NetezzaSQL}"+";SERVER=prodnzo.catmktg.com; PORT=5480; DATABASE="+database_name+";UID={uid};PWD={pwd}".format(uid=uid, pwd=pwd)
        print ('YellowBrick using Windows odbc connection: '+database_name.lower())
        import pyodbc
        conn=pyodbc.connect(connectstring)
        print ('connected to:',database_name)
        #print('connectstring='+connectstring)

    elif path=='Darwin':
        print ('YellowBrick using MAC jdbc connection: '+database_name.lower())
        import jaydebeapi
        conn = jaydebeapi.connect(jclassname = 'com.impossibl.postgres.jdbc.PGDataSource'
                                    , url = 'jdbc:pgsql://10.21.73.70:5432/'+database_name.lower()+'?networkTimeout=3000000&applicationName=YB'
                                    , driver_args = [uid,pwd]
                                   , jars = '/Users/ebird/Netezza Driver/pgjdbc-ng-0.5-complete.jar')
    elif path=='Linux':
        print ('YellowBrick using Linux jdbc connection: '+database_name.lower())
        import jaydebeapi
        conn = jaydebeapi.connect(jclassname = 'com.impossibl.postgres.jdbc.PGDataSource'
                                    , url = 'jdbc:pgsql://10.21.73.70:5432/'+database_name.lower()+'?networkTimeout=3000000&applicationName=YB'
                                    , driver_args = [uid,pwd]
                                    , jars = '/apps_opensource/drivers/jdbc/yellowbrick/pgjdbc-ng-0.5-complete.jar')
    else:
        print ("Invalid System Platform")
    return conn


def SQLExecute(sql,conn):
    import time
    start_time = time.time()
    cursor = conn.cursor()
    print (sql, flush=True)
    print ('SQL executing...', end='', flush=True)
    cursor.execute(sql)
    elapsed_time = time.time() - start_time
    print ('Done. Elapsed: {elapsed}'.format(elapsed=time.strftime("%H:%M:%S", time.gmtime(elapsed_time))), flush=True)
    cursor.close()
    
    
import matplotlib.ticker as mtick

def line_graph(new_sales, old_sales, cat_sales, data_series, banner):
    import matplotlib.pyplot as plt
    fig, ax1 = plt.subplots(figsize=(19,8))
    ax1.plot(date_series, cat_sales, color='grey',dashes=[6, 4])  
    ax1.set_ylabel('Category Sales', color='black')
    ax1.legend(['Category Sales'],bbox_to_anchor=(1.11,0.05), loc='center left')
    ax1.set_ylim(ymin=0)
#     ax1.set_yticklabels(['{:,}'.format(int(x)) for x in ax1.get_yticks().tolist()])


    ax2 = ax1.twinx()
    ax2.plot(date_series, old_sales, color='royalblue')
    ax2.plot(date_series, new_sales, color='darkorange')        #marker='*', markersize=15, label='Blue stars'darkorange
    ax2.set_ylabel('Brand Sales', color='black')
    #ax2.set_ylim(ymin=0)
#     ax2.set_yticklabels(['{:,}'.format(int(x)) for x in ax2.get_yticks().tolist()])
    # ax2.legend(['Old Brand Sales', brand_text+' Shoppers'])
    ax2.legend(['Old Brand Sales', brand_text+' Sales'],bbox_to_anchor=(1.11,0.20), loc='center left')

    
    fmt = '${x:,.0f}'
    tick = mtick.StrMethodFormatter(fmt)
    ax1.yaxis.set_major_formatter(tick) 
    ax2.yaxis.set_major_formatter(tick) 


    plt.rc('axes', titlesize=22) 
    plt.rc('axes', labelsize=22) 
    plt.rc('legend', fontsize=22)
    plt.rc('xtick', labelsize=22) 
    plt.rc('ytick', labelsize=22)  
    plt.xlabel('Date')
    plt.title('Sales Trend',fontsize=30)
    plt.show()
    #1015
#     fig_name = brand_text
    fig.savefig('pic1'+banner+'.png',bbox_inches='tight')
    

def get_all_upcs_byCatgry(conn):
    import pandas as pd
    import numpy as np
    curs = conn.cursor()
    
    curs.execute('''
    create temp table cat_upc_l3 as
    select distinct ti.trade_item_cd as upc, u.grpnum, u.levelflag, u.hierl3cd as hier_cd
        from py1usta1.public.conv_upcs_ttl u
            join trade_item_owner_hierarchy_v tr 
                on u.hierl3cd::varchar = trade_item_hier_l3_cd
            join trade_item_v ti 
                on tr.trade_item_key = ti.trade_item_key  
        where u.levelflag = 3 and u.productname <> 'Milk' and lgl_entity_nbr = 38
    distribute random
    ''')

    curs.execute(''' 
    create temp table cat_upc_l2 as
    select distinct ti.trade_item_cd as upc, u.grpnum, u.levelflag, u.hierl2cd as hier_cd
        from py1usta1.public.conv_upcs_ttl u
            join trade_item_owner_hierarchy_v tr 
                on u.hierl2cd::varchar = trade_item_hier_l2_cd
            join trade_item_v ti 
                on tr.trade_item_key = ti.trade_item_key  
        where u.levelflag = 2 and lgl_entity_nbr = 38
    distribute random
    ''')

    curs.execute(''' 
    create temp table cat_upc_milk as 
    select distinct ti.trade_item_cd as upc, u.grpnum, u.levelflag, u.hierl3cd as hier_cd
        from py1usta1.public.conv_upcs_ttl u
            join trade_item_owner_hierarchy_v tr 
                on u.hierl3cd::varchar = trade_item_hier_l3_cd
            join trade_item_v ti 
                on tr.trade_item_key = ti.trade_item_key  
        where u.productname = 'Milk' and lgl_entity_nbr = 38
    distribute random
    ''')

    curs.execute(''' 
    create temp table conv_upcs as
    select upc,
           grpnum, 
           u.levelflag,
           case 
                when levelflag = 3 then hierl3cd 
                else hierl2cd 
            end as hier_cd 
    from py1usta1.public.conv_upcs_ttl u 
    /*where u.productname <> 'Milk'*/
    distribute random
    ''')

    # removing duplicate UPCs
    curs.execute(''' drop table if exists upc_cat_all''')
    curs.execute(''' 
    create temp table upc_cat_all as 
    select * from cat_upc_l3
        union
    select * from cat_upc_l2
        union 
    select * from cat_upc_milk
        union 
    select * from conv_upcs
    distribute random
    ''')

    upc_cat_all = pd.read_sql('''select * from upc_cat_all''', conn)
    return upc_cat_all

def get_trans_shoprite(ntwk, stop_dt_yago,stop_date,fg_store,conn):
    import pandas as pd
    import numpy as np
    
    curs = conn.cursor()

    # the following improved codes has been testified on 'Shoppers' and compared against with previous version' data
    curs.execute('''
    create temp table shoprite_sales_upc as 
    select u.grpnum
           ,o.ord_event_key
           ,u.upc
           ,cal_sat_wk_ending_dt

           ,sum(case when cal_dt between u.convdate-364 and '%(stop_dt_yago)s' and series = 'Old'
                    then o.purch_amt else 0 
                    end) as old_sales_ly

           ,sum(case 
                   when cal_dt between u.convdate-182 and '%(stop_date)s' and series = 'Old' 
                    then o.purch_amt else 0
                    end) as old_sales_ty

           ,sum(case 
                   when cal_dt between u.convdate-182 and '%(stop_date)s' and series = 'New' 
                    then o.purch_amt else 0 
                    end) as new_sales_ty

           ,sum(case when cal_dt between u.convdate-364 and '%(stop_dt_yago)s' and series = 'Old'
            then o.purch_qty else 0 
            end) as old_units_ly

           ,sum(case 
                   when cal_dt between u.convdate-182 and '%(stop_date)s' and series = 'Old' 
                    then o.purch_qty else 0
                    end) as old_units_ty

           ,sum(case 
                   when cal_dt between u.convdate-182 and '%(stop_date)s' and series = 'New' 
                    then o.purch_qty else 0 
                    end) as new_units_ty

    from ord_trd_itm_fact_ne_v o
        join touchpoint_v on (ord_touchpoint_key = touchpoint_key)
        join trade_item_v using (trade_item_key)
        join date_v on (ord_date_key = date_key)
        join py1usta1.public.conv_upcs_ttl u on (trade_item_cd = upc) 
    where ntwk_id = %(ntwk)s
        and o.cnsmr_id_typ_loy_cd = 'LOYL'
        and cal_dt between u.convdate-364 and '%(stop_date)s' 
        and site_id_txt not in %(fg_store)s
    group by 1, 2, 3, 4
    distribute random'''%locals())


    curs.execute('''
    create temp table shoprite_sales_upc2 as 
    select grpnum
          ,cal_sat_wk_ending_dt
          ,sum(old_sales_ly) as old_sales_ly
          ,sum(old_sales_ty) as old_sales_ty
          ,sum(new_sales_ty) as new_sales_ty

          ,sum(old_units_ly) as old_units_ly
          ,sum(old_units_ty) as old_units_ty
          ,sum(new_units_ty) as new_units_ty
    from shoprite_sales_upc
    group by 1,2
    distribute random'''%locals())

    # ################################################################################################################
    # ############################################ Category Sales ####################################################
    # ################################################################################################################

    #the following improved codes has been testified on 'Shoppers' and compared against with previous version' data
    curs.execute('''
    create temp table shoprite_sales_cat as 
    select u.grpnum
           ,o.ord_event_key
           ,u.upc
           ,cal_sat_wk_ending_dt
           ,sum(case when cal_dt between u.convdate-364 and '%(stop_dt_yago)s' then o.purch_amt else 0 end) as cat_sales_ly
           ,sum(case when cal_dt between u.convdate-182 and '%(stop_date)s'  then o.purch_amt else 0 end) as cat_sales_ty
           ,sum(case when cal_dt between u.convdate-364 and '%(stop_dt_yago)s' then o.purch_qty else 0 end) as cat_units_ly
           ,sum(case when cal_dt between u.convdate-182 and '%(stop_date)s'  then o.purch_qty else 0 end) as cat_units_ty
    from ord_trd_itm_fact_ne_v o
        join touchpoint_v on (ord_touchpoint_key = touchpoint_key)
        join trade_item_v using (trade_item_key) 
        join date_v on (ord_date_key = date_key)
        join py1usta1.public.upc_cat_all_ttl u on (trade_item_cd = upc)
    where ntwk_id = %(ntwk)s
        and o.cnsmr_id_typ_loy_cd = 'LOYL'
        and cal_dt between u.convdate-364 and '%(stop_date)s' 
        and site_id_txt not in %(fg_store)s
    group by 1, 2, 3,4
    distribute random'''%locals())   

    curs.execute('''
    create temp table shoprite_sales_cat2 as 
    select grpnum
          ,cal_sat_wk_ending_dt
          ,sum(cat_sales_ly) as cat_sales_ly
          ,sum(cat_sales_ty) as cat_sales_ty
          ,sum(cat_units_ly) as cat_units_ly
          ,sum(cat_units_ty) as cat_units_ty
    from shoprite_sales_cat
    group by 1,2
    distribute random'''%locals())


    # method 2 --> have to add new upcs (after 11/10/2019) into category sales, not recommend
    #     join py1usta1.public.conv_upcs u 
    #         on case
    #             when u.levelflag = 3 then u.hierl3cd::varchar = trade_item_hier_l3_cd
    #             when u.levelflag = 2 then u.hierl3cd::varchar = trade_item_hier_l2_cd
    #         end

    shoprite_sales_upc = pd.read_sql("select * from shoprite_sales_upc2", conn).set_index('cal_sat_wk_ending_dt')
    shoprite_sales_cat = pd.read_sql("select * from shoprite_sales_cat2", conn).set_index('cal_sat_wk_ending_dt')

    shoprite_sales_cat.index = pd.to_datetime(shoprite_sales_cat.index)
    shoprite_sales_upc.index = pd.to_datetime(shoprite_sales_upc.index)
    
    return(shoprite_sales_cat,shoprite_sales_upc)

def get_trans_pr_fg(ob_not_format_upc, ntwk, stop_dt_yago,stop_date,fg_store,conn):
    import pandas as pd
    import numpy as np
    curs = conn.cursor()

    curs.execute( '''CREATE TEMP TABLE upc_level_all as select
    a.ord_event_key,
    i.trade_item_cd as upc,
    d.cal_sat_wk_ending_dt,
    u.grpnum,
    sum(case when ntwk_id = 301 then a.purch_amt else 0 end) as post_sales_pr,
    sum(case when site_id_txt in %(fg_store)s then a.purch_amt else 0 end) as post_sales_fg,
    sum(case when ntwk_id = 301 then a.purch_qty else 0 end) as post_units_pr,
    sum(case when site_id_txt in %(fg_store)s then a.purch_qty else 0 end) as post_units_fg
    from 
    ord_trd_itm_fact_v a 
        join trade_item_v i on A.TRADE_ITEM_KEY = I.TRADE_ITEM_KEY
        join touchpoint_v t on a.ord_touchpoint_key = t.touchpoint_key
        join date_v d on d.date_key = a.ord_date_key
        join py1usta1.public.upc_cat_all_ttl u on u.upc = i.trade_item_cd
    where 
        (i.trade_item_cd between 4119000000 and 4119099999)
        /*or i.trade_item_cd in %(ob_not_format_upc)s)*/
        and a.purch_qty > 0
        and d.cal_dt >= u.convdate-365
        and d.cal_dt <='%(stop_date)s'
        and lgl_entity_nbr = 38
    group by 1,2,3,4
    order by 1,2,3,4
    distribute random'''%locals())


    curs.execute('''CREATE TEMP TABLE upc_level_series as 
    select a.*, 
        case 
            when b.series = 'New' then 1
            else 0
        end as upc_series
    from upc_level_all a left join py1usta1.public.conv_upcs_ttl b on a.upc = b.upc and a.grpnum = b.grpnum
    distribute random
    '''%locals())


    curs.execute( '''CREATE TEMP TABLE upc_level_series_agg as 
    select grpnum, upc_series, cal_sat_wk_ending_dt
          , sum(post_sales_pr) as post_sales_pr, sum(post_sales_fg) as post_sales_fg
          , sum(post_units_pr) as post_units_pr, sum(post_units_fg) as post_units_fg
    from upc_level_series
    group by 1,2,3
    distribute random
    '''%locals())

    curs.execute('''CREATE TEMP TABLE cat_level_all as select
    a.ord_event_key,
    d.cal_sat_wk_ending_dt,
    u.grpnum,
    sum(case when ntwk_id = 301 then a.purch_amt else 0 end) as post_sales_pr,
    sum(case when site_id_txt in %(fg_store)s then a.purch_amt else 0 end) as post_sales_fg,
    sum(case when ntwk_id = 301 then a.purch_qty else 0 end) as post_units_pr,
    sum(case when site_id_txt in %(fg_store)s then a.purch_qty else 0 end) as post_units_fg
    from 
    ord_trd_itm_fact_v a 
        join trade_item_v i on A.TRADE_ITEM_KEY = I.TRADE_ITEM_KEY
        join touchpoint_v t on a.ord_touchpoint_key = t.touchpoint_key
        join date_v d on d.date_key = a.ord_date_key
        /*inner join trade_item_owner_hierarchy_v as th on (th.trade_item_key = a.trade_item_key)*/
        join py1usta1.public.upc_cat_all_ttl u on u.upc = i.trade_item_cd
    where  a.purch_qty > 0
        and d.cal_dt >= u.convdate-365
        and d.cal_dt <='%(stop_date)s'
        and lgl_entity_nbr = 38
    group by 1,2,3
    order by 1,2,3
    distribute random
    '''%locals())

    curs.execute( '''CREATE TEMP TABLE cat_level_all2 as select
    cal_sat_wk_ending_dt, grpnum
    , sum(post_sales_pr) as post_sales_pr, sum(post_sales_fg) as post_sales_fg
    , sum(post_units_pr) as post_units_pr, sum(post_units_fg) as post_units_fg

    from cat_level_all
    group by 1,2
    order by 1,2
    distribute random
    '''%locals())

    upc_level_all = pd.read_sql('''select * from upc_level_series_agg''',conn).set_index('cal_sat_wk_ending_dt')
    upc_level_all.index = pd.to_datetime(upc_level_all.index)


    cat_level_all = pd.read_sql('''select * from cat_level_all2''',conn).set_index('cal_sat_wk_ending_dt')
    cat_level_all.index = pd.to_datetime(cat_level_all.index)
    
    return(upc_level_all,cat_level_all)
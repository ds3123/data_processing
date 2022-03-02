

import pymysql
import sqlalchemy as sa

# 資料庫設定
engine_2018 = sa.create_engine( 'mysql+pymysql://root:root@localhost:8889/gogopark_2018' ) # gogopark_2018
engine_2021 = sa.create_engine( 'mysql+pymysql://root:root@localhost:8889/gogopark_ts' )   # gogopark_ts






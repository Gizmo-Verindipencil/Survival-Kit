select
    TAB.TABLE_NAME テーブル名
    ,COL.COLUMN_ID カラム番号
    ,case when CON_COL.CONSTRAINT_NAME is not null then '○' else '' end 主キー
    ,COL.COLUMN_NAME カラム名
    ,COL.DATA_TYPE データ型
    ,COL.DATA_LENGTH データ長
    ,COM.COMMENTS コメント
from
    USER_TABLES TAB
left join
    USER_TAB_COLUMNS COL
        on COL.TABLE_NAME = TAB.TABLE_NAME
left join
    USER_CONSTRAINTS CON
        on CON.TABLE_NAME = TAB.TABLE_NAME
left join
    ALL_CONS_COLUMNS CON_COL
        on  CON_COL.TABLE_NAME      = TAB.TABLE_NAME
        and CON_COL.COLUMN_NAME     = COL.COLUMN_NAME
        and CON_COL.CONSTRAINT_NAME = CON.CONSTRAINT_NAME
left join
    USER_COL_COMMENTS COM
        on  COM.TABLE_NAME  = TAB.TABLE_NAME
        and COM.COLUMN_NAME = COL.COLUMN_NAME
where
    CON.CONSTRAINT_TYPE = 'P'
order by
    TAB.TABLE_NAME, COL.COLUMN_ID;
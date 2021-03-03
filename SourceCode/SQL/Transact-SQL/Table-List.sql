SELECT
    テーブル名 = tab.name
    ,カラム番号 = col.column_id
    ,主キー = IIF(pidx.keyno is not null, N'○', N'')
    ,カラム名 = col.name
    ,データ型 = ty.name
    ,データ長 = IIF(col.max_length <> -1, CAST(col.max_length AS varchar), N'max')
    ,コメント = com.value
FROM
    sys.tables AS tab
LEFT JOIN
    sys.columns AS col
        ON tab.object_id = col.object_id
LEFT JOIN
    sys.indexes AS idx
        ON tab.object_id = idx.object_id
LEFT JOIN
    sys.sysindexkeys AS pidx
        ON  tab.object_id = pidx.id
        AND col.column_id = pidx.colid
        AND idx.index_id  = pidx.indid
LEFT JOIN
    sys.types AS ty
        ON  col.system_type_id = ty.system_type_id
        AND col.user_type_id   = ty.user_type_id
LEFT JOIN
    sys.extended_properties AS com
        ON  tab.object_id = com.major_id
        AND col.column_id = com.minor_id
ORDER BY
    tab.name,
    col.column_id;
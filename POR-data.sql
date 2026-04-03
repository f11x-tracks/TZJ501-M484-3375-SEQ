SELECT 
          a0.lot AS lot
         ,To_Char(a0.data_collection_time,'yyyy-mm-dd hh24:mi:ss') AS lot_data_collect_date
         ,a1.entity AS entity
         ,a2.monitor_set_name AS monitor_set_name
         ,a2.test_name AS test_name
         ,a3.measurement_set_name AS measurement_set_name
         ,a10.lo_control_lmt AS lo_control_lmt
         ,a10.up_control_lmt AS up_control_lmt
         ,CASE WHEN REGEXP_INSTR(a5.spc_chart_category,'RESIST=',1,1,1,'i') = 0 THEN NULL ELSE REGEXP_SUBSTR(a5.spc_chart_category,'[^;]+',REGEXP_INSTR(a5.spc_chart_category,'RESIST=',1,1,1,'i'),1,'i') END AS RESIST
         ,a5.valid_flag AS chart_pt_valid_flag
         ,a7.value AS cr_value
         ,a4.wafer AS raw_wafer
         ,a4.WAFER_COORDINATE_X AS X_COORDINATE
         ,a4.WAFER_COORDINATE_Y AS Y_COORDINATE
         ,a4.FOUP_SLOT AS SLOT
FROM 
P_SPC_MEASUREMENT_SET a3
INNER JOIN P_SPC_SESSION a2 ON a2.spcs_id = a3.spcs_id AND a2.data_collection_time=a3.data_collection_time
LEFT JOIN P_SPC_LOT a0 ON a0.spcs_id = a2.spcs_id
INNER JOIN P_SPC_ENTITY a1 ON a2.spcs_id = a1.spcs_id AND a1.entity_sequence=1
INNER JOIN P_SPC_CHART_POINT a5 ON a5.spcs_id = a3.spcs_id AND a5.measurement_set_name = a3.measurement_set_name
LEFT JOIN P_SPC_CHARTPOINT_MEASUREMENT a7 ON a7.spcs_id = a3.spcs_id and a7.measurement_set_name = a3.measurement_set_name
AND a5.spcs_id = a7.spcs_id AND a5.chart_id = a7.chart_id AND a5.chart_point_seq = a7.chart_point_seq AND a5.measurement_set_name = a7.measurement_set_name
LEFT JOIN P_SPC_CHART_LIMIT a10 ON a10.chart_id = a5.chart_id AND a10.limit_id = a5.limit_id
LEFT JOIN P_SPC_MEASUREMENT a4 ON a4.spcs_id = a3.spcs_id AND a4.measurement_set_name = a3.measurement_set_name
AND a4.spcs_id = a7.spcs_id AND a4.measurement_id = a7.measurement_id
WHERE
              a0.data_collection_time >= TRUNC(SYSDATE) - 21 
 AND a3.measurement_set_name Like '%KRF.THICKNESS.MFG%'
 AND a2.TEST_NAME like 'MFGM484%'
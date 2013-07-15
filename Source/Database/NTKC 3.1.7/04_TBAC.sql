prompt PL/SQL Developer import file
prompt Created on Wednesday, July 10, 2013 by QuocAnh
set feedback off
set define off
prompt Loading RCV_DM_TKHAI...
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('04_TBAC', 'M尻 th南g b罪 展烔 ch喈h th南g tin', 'M', '301', to_date('01-01-2009', 'dd-mm-yyyy'), null);
commit;
prompt 1 records loaded
set feedback on
set define on
prompt Done.

prompt PL/SQL Developer import file
prompt Created on Wednesday, July 10, 2013 by QuocAnh
set feedback off
set define off
prompt Loading RCV_GDIEN_TKHAI...
insert into RCV_GDIEN_TKHAI (ID, TEN_CTIEU, COT_01, COT_02, COT_03, COT_04, COT_05, COT_06, COT_07, COT_08, COT_09, COT_10, COT_11, SO_TT, LOAI_DLIEU, MA_CTIEU, COT_12, COT_13, COT_14, COT_15, COT_16, COT_17, COT_18, COT_19, COT_20, COT_21, COT_22)
values (5328, 'M尻 th南g b罪 展烔 ch喈h th南g tin', '1', '2', '3', null, null, null, null, null, null, null, null, 1, '04_TBAC', null, null, null, null, null, null, null, null, null, null, null, null);
commit;
prompt 1 records loaded
set feedback on
set define on
prompt Done.

prompt PL/SQL Developer import file
prompt Created on Wednesday, July 10, 2013 by QuocAnh
set feedback off
set define off
prompt Loading RCV_MAP_CTIEU...
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('04_TBAC', '1', 'C', 5328, '1');
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('04_TBAC', '2', 'C', 5328, '2');
insert into RCV_MAP_CTIEU (LOAI_DLIEU, KY_HIEU, KIEU_DLIEU, GDN_ID, KY_HIEU_CTIEU)
values ('04_TBAC', '3', 'C', 5328, '3');
commit;
prompt 3 records loaded
set feedback on
set define on
prompt Done.

--03/TNDN
insert into QLT_NTK.RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('03_TNDN11', 'Tê khai quyÕt to¸n thuÕ TNDN', 'Y', '330', to_date('14-06-2006', 'dd-mm-yyyy'), null);

--02/TAIN
insert into QLT_NTK.RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('02_TAIN11', 'Tê quyÕt to¸n thuÕ tµi nguyªn', 'Y', '330', to_date('01-07-2011', 'dd-mm-yyyy'), null);

--01/PHLP
insert into QLT_NTK.RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('01_PHLP', 'Tê khai PhÝ, lÖ phÝ (01/PHLP)', 'Y', '330', to_date('01-01-1900', 'dd-mm-yyyy'), null);

--02/PHLP
insert into QLT_NTK.RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('02_PHLP', 'Tê khai PhÝ, lÖ phÝ (02/PHLP)', 'Y', '330', to_date('01-01-1900', 'dd-mm-yyyy'), null);

--03/TD-TAIN
insert into QLT_NTK.RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('03_TD_TAIN', 'Tê khai thuÕ tµi nguyªn(Dµnh cho c¬ së s¶n xuÊt thñy ®iÖn)', 'M', '330', to_date('14-06-2008', 'dd-mm-yyyy'), null);

--02/BVMT
insert into RCV_DM_TKHAI (MA, TEN, KIEU_KY, PHIEN_BAN, START_DATE, END_DATE)
values ('02_BVMT11', 'Tê khai quyÕt to¸n phÝ b¶o vÖ m«i tr­êng', 'Y', '330', to_date('01-07-2011', 'dd-mm-yyyy'), null);

commit;

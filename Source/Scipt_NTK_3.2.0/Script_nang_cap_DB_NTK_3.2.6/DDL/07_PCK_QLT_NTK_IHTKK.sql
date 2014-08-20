CREATE OR REPLACE 
PACKAGE qlt_ntk.ihtkk
  IS
PROCEDURE xnhantthai(
                    noiGui IN VARCHAR2,
                    noiNhan IN VARCHAR2,
                    idTkhai IN VARCHAR2,
                    ngayNhan IN VARCHAR2,
                    tthaiXLy IN VARCHAR2);
PROCEDURE update_TKhai(
                    MVach IN VARCHAR2,
                    id_TKhai IN VARCHAR2);
function  get_us7(data in clob) return clob;
END; -- Package spec
/

-- Grants for Package
GRANT EXECUTE ON qlt_ntk.ihtkk TO dml_role
/
GRANT EXECUTE ON qlt_ntk.ihtkk TO ntk_dml
/

CREATE OR REPLACE 
PACKAGE BODY qlt_ntk.ihtkk
IS
PROCEDURE xnhantthai(
                    noiGui IN VARCHAR2,
                    noiNhan IN VARCHAR2,
                    idTkhai IN VARCHAR2,
                    ngayNhan IN VARCHAR2,
                    tthaiXLy IN VARCHAR2)
IS
    tkhaiCount number(1);
BEGIN
    select count(tkhai_id) into tkhaiCount from ihtkk_ttin_tthai where tkhai_id = idTkhai;
    if(tkhaiCount = 0) then
        insert into ihtkk_ttin_tthai(tkhai_id, noi_gui, noi_nhan, ngay_nhan, tthai) values(idTkhai, noiGui, noiNhan, to_date(ngayNhan, 'DD/MM/RRRR HH24:MI:SS'), tthaiXLy);
    else
        update ihtkk_ttin_tthai set noi_gui=noiGui, noi_nhan=noiNhan, ngay_nhan=to_date(ngayNhan, 'DD/MM/RRRR HH24:MI:SS'), tthai=tthaiXLy where tkhai_id = idTkhai;
    end if;
EXCEPTION
      WHEN others THEN
          rollback;

END;
PROCEDURE update_TKhai(
                    MVach IN VARCHAR2,
                    id_TKhai IN VARCHAR2)
is
v_Temp clob;
p_Clob clob;
strtmp varchar2(32000);
strsql varchar2(32000);
begin_pos number:=0;
end_pos   number:=0;
pos  number:=0;
len number :=0;
charat varchar2(10);
numat number(3);
cursor c_font_map is select chin,chout from IHTKK_FONT_MAP
order by stt;
rec_font_map c_font_map%rowtype;
begin
 strtmp := MVach ;
 open c_font_map;
 loop
   fetch c_font_map into rec_font_map;
   exit when c_font_map%notfound;
   charat := rec_font_map.chin;
   numat :=  rec_font_map.chout;
   strtmp := replace(strtmp,charat,chr(numat));

-- xu ly tieng viet
   --dbms_output.put_line('strtmp:' || substr(strtmp,1,100));

-- append varchar2 to clob
--   strsql := 'insert into tmp values( ''' || strtmp || ''')';
   --execute immediate strsql;
 end loop;

 close c_font_map;
select dlieu_tkhai into p_Clob  from RCV_IHTKK_TKHAI where id = id_TKhai for update ;
 If strtmp  Is Not Null Then
      DBMS_LOB.CREATETEMPORARY(v_Temp
                              ,TRUE);
      DBMS_LOB.WRITEAPPEND(v_Temp
                          ,Length(strtmp)
                          ,strtmp);

      IF DBMS_LOB.Substr(p_Clob) is Not Null Then
        Dbms_Lob.Append(p_Clob
                       ,v_Temp);
      Else
        p_Clob := v_Temp;
      End IF;

      DBMS_LOB.FREETEMPORARY(v_Temp);
    End If;
update RCV_IHTKK_TKHAI set dlieu_tkhai = p_Clob  where id = id_TKhai ;
/*
strsql := 'update RCV_IHTKK_MVACH set dlieu_mvach = dlieu_mvach || '''|| MVach || ''' where id = ' || id_TKhai ;
--dbms_output.put_line('strsql' || substr(strsql,1,100));
--insert into tmp values(strsql);

--update RCV_IHTKK_MVACH set dlieu_mvach = p_Clob  where id = id_TKhai ;
execute immediate strsql ;*/
--commit;
end;
function  get_us7(data in clob) return clob
is
tmp varchar2(32767);
beginPos number;
pos number :=1;
len number:=0;
tmp_lob clob;
begin
--execute immediate msg;
len := dbms_lob.getlength(data);
while (pos < len) loop
  pos := dbms_lob.instr(data,'chr(',pos,1);

end loop;
return tmp_lob;
end;
END;
/


-- End of DDL Script for Package Body QLT_NTK.IHTKK


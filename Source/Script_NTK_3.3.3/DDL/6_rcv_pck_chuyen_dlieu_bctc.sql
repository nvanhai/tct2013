CREATE OR REPLACE PACKAGE BODY QLT_NTK.rcv_pck_chuyen_dlieu_bctc IS
    Procedure Prc_load_ntk_2_bctc_log is
    errm varchar2(2000);
    v_phongxl varchar2(2000);
    v_ma_cqt varchar2(10);
    v_so_hs_goc varchar2(100);
      BEGIN
        For rec in (select id,itkhai_id,tin,ten_dtnt,dia_chi,ngay_cap_nhat,nguoi_cap_nhat,ngay_nop,phong_xly,kkbs,Kykk_Tu_Ngay, hthuc_nop  from rcv_tkhai_hdr hdr where rownum <=500 and loai_tkhai in ('15_BCTC','15_BCTC10','48_BCTC','48_BCTC13','95_BCTC','16_BCTC') and (da_nhan is null or da_nhan='N') order by tin,kkbs,ngay_nop)
          loop
            begin
              /*select trim(gia_tri)||'01' as phongxl into v_phongxl from qlt_owner.qlt_tham_so where ten='MA_CQT';*/
              savepoint exec_tran;
              select ma_cqt||'01' into v_phongxl from rcv_v_dtnt where tin=rec.tin;
              -- insert vao rcv_tkhai_hdr_bctc_log;
              insert into rcv_tkhai_hdr_bctc_log (ID,TIN,LOAI_TKHAI,NGAY_NOP,KYKK_TU_NGAY,NGAY_CAP_NHAT,NGUOI_CAP_NHAT,CO_LOI_DDANH,SO_TT_TK,PHONG_XLY,KKBS)
              select ID,TIN,LOAI_TKHAI,NGAY_NOP,KYKK_TU_NGAY,NGAY_CAP_NHAT,NGUOI_CAP_NHAT,CO_LOI_DDANH,SO_TT_TK,v_phongxl,KKBS from rcv_tkhai_hdr
              where id=rec.id;
              -- insert vao rcv_tkhai_dtl_bctc_log;
              insert into rcv_tkhai_dtl_bctc_log (ID,HDR_ID,LOAI_DLIEU,KY_HIEU,GIA_TRI)
              select  ID,HDR_ID,LOAI_DLIEU,KY_HIEU,GIA_TRI from rcv_tkhai_dtl where hdr_id=rec.id;
              update rcv_tkhai_hdr a set a.da_nhan='Y' where a.id=rec.id;
              insert into qhs_so_hoso (id,dhs_ma,so_hoso,tin,ten,dia_chi,ngay_nhan,ngay_nhap,nguoi_nhap,phong_xly,ngay_nop,tthai_hoso,so_hoso_goc,kykk_tu_ngay, hthuc_nop)
              values (qhs_so_hoso_seq.nextval,'39',rcv_pck_chuyen_dlieu_bctc.Fnc_Get_So_Hoso(REPLACE(TO_CHAR(SYSDATE,'YY/MM/DD'),'/'),'39'),rec.tin,rec.ten_dtnt,rec.dia_chi,rec.ngay_cap_nhat,sysdate,rec.nguoi_cap_nhat,rec.phong_xly,rec.ngay_nop,DECODE(rec.kkbs,0,'01','03'),rcv_pck_chuyen_dlieu_bctc.Fnc_So_Hoso_exist(rec.tin,rec.kykk_tu_ngay),rec.kykk_tu_ngay, decode(rec.hthuc_nop,'I','01','02'));
              select trim(gia_tri) into v_ma_cqt from qlt_owner.qlt_tham_so where ten='MA_CQT';
              if(rec.itkhai_id is not null) then
               insert into ihtkk_ttin_tthai (Tkhai_Id,Noi_Gui,Noi_Nhan,Ngay_Nhan,Tthai) values (rec.itkhai_id,v_ma_cqt,'00000',sysdate,'24');
              end if;
              commit;
              EXCEPTION
                WHEN OTHERS THEN
                  begin
                    errm := SQLERRM;
                    ROLLBACK TO exec_tran;
                    update rcv_tkhai_hdr set da_nhan='E' where id=rec.id;
                    --Update cho IHTKK
                    if(rec.itkhai_id is not null) then
                     insert into iHTKK_TTIN_TTHAI (Tkhai_Id,Noi_Gui,Noi_Nhan,Ngay_Nhan,Tthai) values(rec.itkhai_id,v_ma_cqt,'00000',sysdate,'25');
                    end if;
                    commit;
                  end;
            end;
          end loop;
       END;
   FUNCTION Fnc_Get_So_Hoso(p_Date IN VARCHAR2, p_Ma IN VARCHAR2) RETURN VARCHAR2 IS
    v_so_hs_cuoi     NUMBER;
    v_So_Hoso VARCHAR2(30);
    v_Tmp VARCHAR2(50);
    --Lay so ho so lon nhat(phan sau /)
    v_Num  NUMBER(20) := 0;
  BEGIN
    v_Tmp  := p_Date || '/' || p_Ma || '/';
    SELECT MAX(TO_NUMBER(SUBSTR(so_hoso, INSTR(so_hoso, '/', -1, 1) + 1 )+1 ))
    INTO v_so_hs_cuoi
    FROM qhs_so_hoso hs
    WHERE SUBSTR(hs.so_hoso,1,2)=TO_CHAR(SYSDATE,'YY');
    IF (v_so_hs_cuoi IS NULL) THEN
      v_so_hs_cuoi := 1;
    END IF;
    v_So_Hoso := v_Tmp || TO_CHAR(v_so_hs_cuoi);
    RETURN v_So_Hoso;
  END;
  FUNCTION Fnc_So_Hoso_exist(p_tin IN VARCHAR2, p_kykk_tu_ngay IN DATE) RETURN VARCHAR2 IS
  v_tong_so_hoso number;
  v_so_hoso varchar2(50);
  v_tmp varchar2(50);
  BEGIN
    SELECT count(*)
    INTO v_tong_so_hoso
      FROM qhs_so_hoso hso, qhs_dm_hoso dm
      WHERE (dm.ma = hso.dhs_ma)
      AND (dm.loai = '01')
      AND (hso.kykk_tu_ngay = p_kykk_tu_ngay)
      AND (hso.dhs_ma = '39')
      AND (hso.tin = p_tin);
    if v_tong_so_hoso=0 then v_so_hoso:=null;
      else
         begin
            SELECT so_hoso
            INTO v_tmp
              FROM qhs_so_hoso hso, qhs_dm_hoso dm
              WHERE (dm.ma = hso.dhs_ma)
              AND (dm.loai = '01')
              AND (hso.kykk_tu_ngay = p_kykk_tu_ngay)
              AND (hso.dhs_ma = '39')
              AND (hso.tin = p_tin)
              AND TTHAI_HOSO='01';
             v_so_hoso:=v_tmp;
         end;
    end if;
    Return v_so_hoso;
  END;
END;

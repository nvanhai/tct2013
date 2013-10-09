-- Start of DDL Script for Package Body QLT_OWNER.RCV_PCK_CHUYEN_DLIEU_QLT
-- Generated 8-Dec-2005 18:51:44 from QLT_OWNER@QLT_93

CREATE OR REPLACE 
PACKAGE rcv_pck_chuyen_dlieu_qlt IS
/*******************************************************************************
Phien ban: 1.0
Nguoi lap: Nguyen Ta Anh
Ngay lap: 15/11/2005
Muc dich: Do du lieu tu CSDL trung gian sau khi quet to khai vao QLT
*******************************************************************************/
  TYPE Record_Hdr IS RECORD(id NUMBER(10,0),
                            tin VARCHAR2(14),
                            ten_dtnt VARCHAR2(60),
                            dia_chi VARCHAR2(60),
                            loai_tkhai VARCHAR2(2),
                            ngay_nop DATE,
                            kylb_tu_ngay DATE,
                            kylb_den_ngay DATE,
                            kykk_tu_ngay DATE,
                            kykk_den_ngay DATE,
                            ngay_cap_nhat DATE,
                            nguoi_cap_nhat VARCHAR2(60),
                            co_loi_ddanh CHAR(1),
                            so_hieu_tep VARCHAR(20),
                            so_tt_tk NUMBER(10,0),
                            ghi_chu_loi VARCHAR(100),
                            co_gtrinh_02a CHAR(1),
                            co_gtrinh_02b CHAR(1),
                            co_gtrinh_02c CHAR(1));
  PROCEDURE Prc_Chuyen_Dlieu_Qlt;
/******************************************************************************/
  FUNCTION Fnc_TKhai_Exits(p_Tin VARCHAR2,
                           p_Loai_TKhai VARCHAR2,
                           p_KyKK DATE,
                           p_Tthai OUT VARCHAR2) RETURN NUMBER;
/******************************************************************************/
  FUNCTION Fnc_Lay_So_Ktru_Ktruoc(p_Tin VARCHAR2,
                                  p_Kylb DATE) RETURN NUMBER;
/******************************************************************************/
  PROCEDURE Prc_Insert_Header_Detail(p_Record_Of_Header Record_Hdr);
/******************************************************************************/
  FUNCTION Fnc_Ktra_Ctieu(p_So_CQT NUMBER,
                          p_So_DTNT NUMBER,
                          p_Kylb DATE) RETURN BOOLEAN;
/******************************************************************************/
END;
/

-- Grants for Package
GRANT EXECUTE ON rcv_pck_chuyen_dlieu_qlt TO qlt
/

CREATE OR REPLACE 
PACKAGE BODY rcv_pck_chuyen_dlieu_qlt IS
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/11/2005
Muc dich: Thuc hien kiem tra to khai da ton tai trong QLT chua,
          ung voi mot ky ke khai
Tham so:
     - p_Tin: Ma so thue cua DTNT
     - p_Loai_TKhai: Loai to khai cua DTNT nop
     - p_KyKK: Ky ke khai ma DTNT nop to khai
*******************************************************************************/
    FUNCTION Fnc_TKhai_Exits(p_Tin VARCHAR2,
                             p_Loai_TKhai VARCHAR2,
                             p_KyKK DATE,
                             p_Tthai OUT VARCHAR2) RETURN NUMBER IS
        CURSOR c_TKhai_Exits IS
            SELECT hdr.id, hdr.tthai
            FROM qlt_tkhai_hdr hdr
            WHERE (hdr.tin = p_Tin)
              AND (hdr.dtk_ma_loai_tkhai = p_Loai_TKhai)
              AND (hdr.kykk_tu_ngay = p_KyKK)
              AND (hdr.ltd = 0);
        vc_TKhai_Exits c_TKhai_Exits%ROWTYPE;
    BEGIN
        OPEN c_TKhai_Exits;
        FETCH c_TKhai_Exits INTO vc_TKhai_Exits;
        IF (c_TKhai_Exits%FOUND) THEN
            p_Tthai := vc_TKhai_Exits.tthai;
            RETURN vc_TKhai_Exits.id;
        ELSE
            RETURN NULL;
        END IF;
        CLOSE c_TKhai_Exits;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/11/2005
Muc dich: Thuc hien lay so khau tru ky truoc (trong so thue) cua ky ke khai
          hien tai
Tham so: 
        - p_Tin: Ma so thue cua DTNT
        - p_Kylb: Ky lap bo cua co quan thue
*******************************************************************************/
    FUNCTION Fnc_Lay_So_Ktru_Ktruoc(p_Tin VARCHAR2,
                                    p_Kylb DATE) RETURN NUMBER IS
        v_Ktru NUMBER(20,2);
    BEGIN
        SELECT SUM(ktru_dky) INTO v_Ktru
		FROM qlt_so_thue
		WHERE (tin = p_Tin)
		  AND (kylb_tu_ngay = TRUNC(p_Kylb,'MONTH'))
		  AND (kylb_den_ngay = LAST_DAY(p_Kylb))
		  AND (tmt_ma_muc = '014')
		  AND (tmt_ma_tmuc = '01');
        IF (v_Ktru < 0) THEN
            v_Ktru := 0;
        END IF;
		RETURN ROUND(NVL(v_Ktru,0));
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 17/11/2005
Muc dich: Thuc hien kiem tra dung sai theo nguong (ve mat so hoc) 2 chi tieu
        - Hang hoa, dich vu ban ra chiu thue suat 5%
        - Hang hoa, dich vu ban ra chiu thue suat 10%
Tham so:
        - p_So_CQT: So thue cua CQT tinh
        - p_So_DTNT: So thue cua DTNT tinh
        - p_Kylb: Ky lap bo cua CQT
*******************************************************************************/
    FUNCTION Fnc_Ktra_Ctieu(p_So_CQT NUMBER,
                            p_So_DTNT NUMBER,
                            p_Kylb DATE) RETURN BOOLEAN IS
        CURSOR c_Gia_Tri(p_Loai VARCHAR2) IS
        SELECT NVL(gia_tri,0) gia_tri
        FROM qlt_dkien_ktra_tkhai
        WHERE (lth_ma_thue = '01')
          AND (loai_dkien = p_Loai)
          AND (tu_ky <= p_Kylb)
          AND (NVL(den_ky, TO_DATE('16/11/2050','DD/MM/RRRR')) >= p_Kylb);

        vc_Gia_Tri c_Gia_Tri%ROWTYPE;
        v_Tyle BOOLEAN;
        v_GTTD BOOLEAN;
        v_Gt_Tyle NUMBER(20,2);
        v_Gt_GTTD NUMBER(20,2);
        v_Kiem_Tra NUMBER(20,2);
    BEGIN
    	IF (p_So_DTNT = p_So_CQT) THEN
    		RETURN TRUE;
    	END IF;
    	
    	OPEN c_Gia_Tri('0');
    	FETCH c_Gia_Tri INTO vc_Gia_Tri;
    	IF (c_Gia_Tri%NOTFOUND) THEN
    	    CLOSE c_Gia_Tri;
    	    v_Tyle := FALSE;
    	ELSE
    	    v_Gt_Tyle := vc_Gia_Tri.gia_tri;
    	    v_Tyle := TRUE;
    	    CLOSE c_Gia_Tri;
    	END IF;
    	
    	OPEN c_Gia_Tri('1');
    	FETCH c_Gia_Tri INTO vc_Gia_Tri;
        IF c_Gia_Tri%NOTFOUND THEN
            CLOSE c_Gia_Tri;
            v_GTTD := FALSE;
        ELSE
            v_Gt_GTTD := vc_Gia_Tri.gia_tri;
            CLOSE c_Gia_Tri;
            v_GTTD := TRUE;
        END IF;

        IF (v_Tyle AND v_GTTD) THEN --Co ca 2 tham so
            v_Kiem_Tra := GREATEST(p_So_DTNT, p_So_CQT);
            IF (v_Kiem_Tra <> 0) THEN
        		IF (ABS(p_So_DTNT - p_So_CQT)*100/v_Kiem_Tra > v_Gt_Tyle) THEN
        			RETURN FALSE;
        		END IF;
            END IF;

        	IF (ABS(p_So_DTNT - p_So_CQT) > v_Gt_GTTD) THEN
        		RETURN FALSE;
        	END IF;
        	RETURN TRUE;
        END IF;

        IF (v_Tyle AND NOT(v_GTTD)) THEN -- Chi co tham so ty le
            v_Kiem_Tra := GREATEST(p_So_DTNT, p_So_CQT);
            IF (v_Kiem_Tra <> 0) THEN
            	IF (ABS(p_So_DTNT - p_So_CQT)*100/v_Kiem_Tra > v_Gt_Tyle) THEN
            		RETURN FALSE;
            	ELSE
            		RETURN TRUE;
            	END IF;
           	ELSE
            	RETURN TRUE;				
            END IF;
        END IF;

        IF (v_Tyle = FALSE AND v_GTTD = TRUE) THEN -- Chi co tham so GTTD
        	IF (ABS(p_So_DTNT - p_So_CQT) > v_Gt_GTTD) THEN
        		RETURN FALSE;
        	ELSE
        		RETURN TRUE;
        	END IF;
        END IF;

        IF (v_Tyle = FALSE AND v_GTTD = FALSE) THEN -- Khong co tham so nao
            IF (p_So_DTNT <> p_So_CQT) THEN
                RETURN FALSE;
           	ELSE
            	RETURN TRUE;
            END IF;
        END IF;
        RETURN TRUE;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 17/11/2005
Muc dich: Thuc hien do du lieu to khai vao cac bang QLT_TKHAI_HDR,
          QLT_TKHAI_GTGT_KT, QLT_GTRINH_GTGT_KT_02A,
          QLT_GTRINH_GTGT_KT_02BC, QLT_XLTK_GDICH
Tham so: 
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
*******************************************************************************/
    PROCEDURE Prc_Insert_Header_Detail(p_Record_Of_Header Record_Hdr) IS
	    TYPE Record_Dtl IS RECORD(id NUMBER(10,0),
                                  ctk_id NUMBER(10,0),
                                  so_tt NUMBER(3,0),
                                  doanhso_dtnt NUMBER(20,2),
                                  sothue_dtnt NUMBER(20,2),
                                  doanhso_cqt NUMBER(20,2),
                                  sothue_cqt NUMBER(20,2),
                                  ke_khai_sai VARCHAR2(1),
                                  ma_so_ct_thue VARCHAR2(5),
                                  ma_so_ct_doanhso VARCHAR2(5));
        TYPE Array_Of_Record_Dtl IS TABLE OF Record_Dtl INDEX BY BINARY_INTEGER;
        v_Array_Of_Record_Dtl Array_Of_Record_Dtl;

        CURSOR c_Dtnt IS
            SELECT pay_taxo_id ma_cqt,
                   tran_prov ma_tinh,
                   tran_dist ma_huyen,
                   depa_id ma_phong,
                   staff_id ma_canbo,
                   level_code cap,
                   category chuong,
                   group_code loai,
                   chapter khoan
            FROM tin_v_payer
            WHERE (tin = p_Record_Of_Header.tin)
              AND (update_no = 0);
        vc_Dtnt c_Dtnt%ROWTYPE;
        
        CURSOR c_TKhai_Dtl IS
            SELECT tkhai.ctk_id,
                   tkhai.so_tt,
                   DECODE(tkhai.ctk_id, 252,DECODE(tkhai.doanhso_dtnt, 'x', 1, 0), tkhai.doanhso_dtnt) doanhso_dtnt,
                   tkhai.sothue_dtnt,
                   tkhai.ky_hieu_ctieu_ds,
                   tkhai.ky_hieu_ctieu_st
            FROM rcv_v_tkhai_gtgt_kt tkhai
            WHERE (tkhai.hdr_id = p_Record_Of_Header.id)
            ORDER BY tkhai.ctk_id;
        vc_TKhai_Dtl c_TKhai_Dtl%ROWTYPE;
        
        CURSOR c_PLuc2B IS
            SELECT ctg_id,
                   gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2b;
        vc_PLuc2B c_PLuc2B%ROWTYPE;
        
        CURSOR c_PLuc2C IS
            SELECT ctg_id,
                   gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2c;
        vc_PLuc2C c_PLuc2C%ROWTYPE;
        
        CURSOR c_AnDinh_Exist IS
            SELECT 1
            FROM qlt_ds_an_dinh_hdr hdr
            WHERE (hdr.tin = p_Record_Of_Header.tin)
              AND (hdr.dtk_ma = p_Record_Of_Header.loai_tkhai)
              AND (hdr.kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay);
        vc_AnDinh_Exist c_AnDinh_Exist%ROWTYPE;
        
        v_Thue_Dau_Ra NUMBER(20,2);
        v_Thue_Dau_Vao NUMBER(20,2);
        v_Thue_KTru_KySau NUMBER(20,2);
        v_Thue_PSinh_KyNay NUMBER(20,2);
        v_Thue_KTru_KyNay NUMBER(20,2);
        v_Thue_PNop_KyNay NUMBER(20,2);
        v_Hdr_Id NUMBER(10,0);
        v_TKhai_Exits_Id NUMBER;
        v_So_KTru NUMBER(20,2);
        v_PSinh_TKy NUMBER(20,2);
        v_Tthai_Tkhai VARCHAR2(1);
        v_Count NUMBER(10) := 0;
        v_Index NUMBER(10) := 0;
        
    BEGIN
        /*Lay thong tin DTNT tu CSDL TIN*/
        OPEN c_Dtnt;
        FETCH c_Dtnt INTO vc_Dtnt;
        IF (c_Dtnt%NOTFOUND) THEN
            RETURN;
        END IF;
        CLOSE c_Dtnt;

        /*Xu ly voi du lieu to khai Detail*/
        FOR vc_TKhai_Dtl IN c_TKhai_Dtl LOOP
            /*Dua du lieu tu RCV_V_TKHAI_GTGT_KT ra mang de xu ly*/
            v_Count := v_Count + 1;
            SELECT qlt_xltk_dtl_seq.NEXTVAL INTO v_Array_Of_Record_Dtl(v_count).id
            FROM dual;
            v_Array_Of_Record_Dtl(v_Count).ctk_id := vc_TKhai_Dtl.ctk_id;
            v_Array_Of_Record_Dtl(v_Count).so_tt := vc_TKhai_Dtl.so_tt;
            v_Array_Of_Record_Dtl(v_Count).doanhso_dtnt := vc_TKhai_Dtl.doanhso_dtnt;
            v_Array_Of_Record_Dtl(v_Count).sothue_dtnt := vc_TKhai_Dtl.sothue_dtnt;
            v_Array_Of_Record_Dtl(v_Count).doanhso_cqt := vc_TKhai_Dtl.doanhso_dtnt;
            v_Array_Of_Record_Dtl(v_Count).sothue_cqt := vc_TKhai_Dtl.sothue_dtnt;
            v_Array_Of_Record_Dtl(v_Count).ke_khai_sai := NULL;
            v_Array_Of_Record_Dtl(v_Count).ma_so_ct_thue := vc_TKhai_Dtl.ky_hieu_ctieu_st;
            v_Array_Of_Record_Dtl(v_Count).ma_so_ct_doanhso := vc_TKhai_Dtl.ky_hieu_ctieu_ds;
        END LOOP;

        /*Thuc hien kiem tra ke khai sai voi cac chi tieu to khai*/
        /*********************************************************/

        /*Lay so khau ky truoc cua co quan thue, chi tieu 11*/
        v_Array_Of_Record_Dtl(2).sothue_cqt := Fnc_Lay_So_Ktru_Ktruoc(p_Record_Of_Header.tin,
                                                                      p_Record_Of_Header.kylb_tu_ngay);

        /* Xu ly chi tieu 12 = 14 + 16 */
        v_Array_Of_Record_Dtl(5).doanhso_cqt := v_Array_Of_Record_Dtl(6).doanhso_cqt +
                                                 v_Array_Of_Record_Dtl(7).doanhso_cqt;

        /* Xu ly chi tieu 13 = 15 + 17 */
        v_Array_Of_Record_Dtl(5).sothue_cqt := v_Array_Of_Record_Dtl(6).sothue_cqt +
                                               v_Array_Of_Record_Dtl(7).sothue_cqt;

        /* Xu ly chi tieu 22 = 13 + 19 - 21 */
        v_Array_Of_Record_Dtl(11).sothue_cqt := v_Array_Of_Record_Dtl(5).sothue_cqt +
                                                v_Array_Of_Record_Dtl(9).sothue_cqt -
                                                v_Array_Of_Record_Dtl(10).sothue_cqt;

        /* Xu ly chi tieu 27 = 29 + 30 + 32 */
        v_Array_Of_Record_Dtl(17).doanhso_cqt := v_Array_Of_Record_Dtl(17).doanhso_cqt +
                                                 v_Array_Of_Record_Dtl(18).doanhso_cqt +
                                                 v_Array_Of_Record_Dtl(19).doanhso_cqt;

        /* Xu ly chi tieu 24 = 26 + 27 */
        v_Array_Of_Record_Dtl(14).doanhso_cqt := v_Array_Of_Record_Dtl(15).doanhso_cqt +
                                                 v_Array_Of_Record_Dtl(17).doanhso_cqt;

        /* Xu ly chi tieu 28 = 31 + 33 */
        v_Array_Of_Record_Dtl(16).sothue_cqt := v_Array_Of_Record_Dtl(18).sothue_cqt +
                                                v_Array_Of_Record_Dtl(19).sothue_cqt;

        /* Xu ly chi tieu 25 = 28*/
        v_Array_Of_Record_Dtl(14).sothue_cqt := v_Array_Of_Record_Dtl(16).sothue_cqt;

        /* Xu ly chi tieu 38 = 24 + 34 - 36 */
        v_Array_Of_Record_Dtl(23).doanhso_cqt := v_Array_Of_Record_Dtl(14).doanhso_cqt +
                                                 v_Array_Of_Record_Dtl(21).doanhso_cqt -
                                                 v_Array_Of_Record_Dtl(22).doanhso_cqt;
        /* Xu ly chi tieu 39 = 25 + 35 - 37 */
        v_Array_Of_Record_Dtl(23).sothue_cqt := v_Array_Of_Record_Dtl(14).sothue_cqt +
                                                 v_Array_Of_Record_Dtl(21).sothue_cqt -
                                                 v_Array_Of_Record_Dtl(22).sothue_cqt;

        /* Xu ly chi tieu 40 = 39 - 23 - 11 */
        v_Array_Of_Record_Dtl(25).sothue_cqt := v_Array_Of_Record_Dtl(23).sothue_cqt -
                                                v_Array_Of_Record_Dtl(12).sothue_dtnt -
                                                v_Array_Of_Record_Dtl(2).sothue_cqt;
        /* Xu ly chi tieu 41 = |40| */
        v_Array_Of_Record_Dtl(26).sothue_cqt := ABS(v_Array_Of_Record_Dtl(25).sothue_cqt);

        /*Xu ly chi tieu "Thue GTGT con duoc khau tru chuyen ky sau", chi tieu 43*/
        v_Array_Of_Record_Dtl(28).sothue_cqt := v_Array_Of_Record_Dtl(26).sothue_cqt -
                                                v_Array_Of_Record_Dtl(27).sothue_cqt;

        FOR i IN 1..v_Count LOOP
            IF (i = 18) THEN --Chi tieu 31
                IF (NOT Fnc_Ktra_Ctieu(ROUND((v_Array_Of_Record_Dtl(18).doanhso_cqt*5)/100),
                    v_Array_Of_Record_Dtl(18).sothue_cqt,p_Record_Of_Header.kylb_tu_ngay))
                THEN
                    v_Array_Of_Record_Dtl(18).ke_khai_sai := 'Y';
                    v_Array_Of_Record_Dtl(18).sothue_cqt :=
                               ROUND((v_Array_Of_Record_Dtl(18).doanhso_cqt*5)/100);
                END IF;
            ELSIF (i = 19) THEN --Chi tieu 33
                IF (NOT Fnc_Ktra_Ctieu(ROUND((v_Array_Of_Record_Dtl(19).doanhso_cqt*5)/100),
                    v_Array_Of_Record_Dtl(19).sothue_cqt,p_Record_Of_Header.kylb_tu_ngay))
                THEN
                    v_Array_Of_Record_Dtl(19).ke_khai_sai := 'Y';
                    v_Array_Of_Record_Dtl(19).sothue_cqt :=
                               ROUND((v_Array_Of_Record_Dtl(18).doanhso_cqt*5)/100);
                END IF;
            ELSIF (i = 27) THEN --Chi tieu 42 so sanh voi 41
                IF (v_Array_Of_Record_Dtl(27).sothue_cqt >
                    v_Array_Of_Record_Dtl(26).sothue_cqt)
                THEN
                    v_Array_Of_Record_Dtl(27).ke_khai_sai := 'Y';
                    v_Array_Of_Record_Dtl(27).sothue_cqt :=
                               v_Array_Of_Record_Dtl(26).sothue_cqt;
                END IF;
            ELSE
                IF (v_Array_Of_Record_Dtl(i).doanhso_dtnt <>
                    v_Array_Of_Record_Dtl(i).doanhso_cqt
                    OR
                    v_Array_Of_Record_Dtl(i).sothue_dtnt <>
                    v_Array_Of_Record_Dtl(i).sothue_cqt)
                THEN
                    v_Array_Of_Record_Dtl(i).ke_khai_sai := 'Y';
                END IF;
            END IF;
        END LOOP;

        /*Tinh toan so thue phat sinh*/
        v_Thue_Dau_Ra := v_Array_Of_Record_Dtl(23).sothue_dtnt; --Chi tieu 39
        v_Thue_Dau_Vao := v_Array_Of_Record_Dtl(12).sothue_dtnt; --Chi tieu 23
        v_Thue_KTru_KySau := v_Array_Of_Record_Dtl(28).sothue_dtnt; --Chi tieu 43;
        v_Thue_PSinh_KyNay := v_Array_Of_Record_Dtl(23).sothue_dtnt -
                              v_Array_Of_Record_Dtl(12).sothue_dtnt; -- = 39-23
        v_Thue_KTru_KyNay := v_Array_Of_Record_Dtl(2).sothue_dtnt +
                             v_Array_Of_Record_Dtl(12).sothue_dtnt -
                             v_Array_Of_Record_Dtl(27).sothue_dtnt -
                             v_Array_Of_Record_Dtl(28).sothue_dtnt; -- = 11+23-42-43
        v_Thue_PNop_KyNay := v_Array_Of_Record_Dtl(25).sothue_dtnt; --Chi tieu 40;
        /*Ket thuc tinh toan so thue phat sinh*/

        /********************************************************/
        /*Ket thuc kiem tra ke khai sai voi cac chi tieu to khai*/

        /*Kiem tra to khai da ton tai trong ky ke khai chua*/
        v_TKhai_Exits_Id := Fnc_TKhai_Exits(p_Record_Of_Header.tin,
                                            p_Record_Of_Header.loai_tkhai,
                                            p_Record_Of_Header.kykk_tu_ngay,
                                            v_Tthai_Tkhai);

        /*Xu ly to khai*/
        IF (v_TKhai_Exits_Id IS NOT NULL) THEN
            /*Neu da ton tai to khai trong ky ke khai*/

            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Sinh giao dich*/
            Qlt_Pck_Gdich.Prc_Set_GTGT_2004;

            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR','14',v_TKhai_Exits_Id);

            /*Thong tin Header*/
            UPDATE qlt_tkhai_hdr
               SET co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   co_gtrinh_02a = p_Record_Of_Header.co_gtrinh_02a,
                   co_gtrinh_02b = p_Record_Of_Header.co_gtrinh_02b,
                   co_gtrinh_02c = p_Record_Of_Header.co_gtrinh_02c,
                   so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = v_Tthai_Tkhai, --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat
            WHERE (id = v_TKhai_Exits_Id)
              AND (ltd = 0);

            /*Thong tin Detail*/
            FOR id IN 1..v_Count LOOP
                v_Index := v_Index + 1;
                UPDATE qlt_tkhai_gtgt_kt
                   SET doanhso_dtnt = v_Array_Of_Record_Dtl(v_Index).doanhso_dtnt,
                       sothue_dtnt = v_Array_Of_Record_Dtl(v_Index).sothue_dtnt,
                       doanhso_cqt = NVL(v_Array_Of_Record_Dtl(v_Index).doanhso_cqt,0),
                       sothue_cqt = NVL(v_Array_Of_Record_Dtl(v_Index).sothue_cqt,0),
                       ke_khai_sai = v_Array_Of_Record_Dtl(v_Index).ke_khai_sai
                WHERE (tkh_id = v_TKhai_Exits_Id)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_Array_Of_Record_Dtl(v_Index).ctk_id);
            END LOOP;

            /*Thong tin bang phat sinh*/
            UPDATE qlt_htoan_tkhai_gtgt_kt_2004
               SET thue_psinh_knay = v_Thue_PSinh_KyNay,
                   thue_ktru_knay = v_Thue_KTru_KyNay,
                   thue_pnop_knay = v_Thue_PNop_KyNay,
                   thue_ktru_ksau = v_Thue_KTru_KySau,
                   thue_dau_vao = v_Thue_Dau_Vao,
                   thue_dau_ra = v_Thue_Dau_Ra
            WHERE (tkh_id = v_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

            /*Thong tin phu luc 2A*/
            /*Xoa ban ghi cu*/
            DELETE FROM qlt_gtrinh_gtgt_kt_02a
            WHERE (tkh_id = v_TKhai_Exits_Id)
              AND (tkh_ltd = 0);
            /*Insert ban ghi moi*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02a(id,
                                               tkh_id,
                                               tkh_ltd,
                                               ctk_id,
                                               dien_giai,
                                               ky_hieu,
                                               kykk_tu_ngay,
                                               kykk_den_ngay,
                                               gia_tri_kkhai,
                                               gia_tri_dchinh,
                                               gia_tri_hhdv,
                                               gia_tri_thue,
                                               ghi_chu)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   v_TKhai_Exits_Id,
                   0,
                   DECODE(pluc2a.ky_hieu,'18',255,'19',255,'20',255,'21',255,
                                         '34',264,'35',264,'36',264,'37',264) ctk_id,
                   pluc2a.dien_giai,
                   pluc2a.ky_hieu,
                   TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH'),
                   LAST_DAY(TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH')),
                   pluc2a.gia_tri_slieu_kkhai,
                   pluc2a.gia_tri_slieu_dchinh,
                   pluc2a.gia_tri_hhdv,
                   pluc2a.gia_tri_thue_gtgt,
                   pluc2a.gia_tri_lydo_dchinh
            FROM rcv_v_tkhai_gtgt_kt_pluc2a pluc2a
            WHERE (pluc2a.hdr_id = p_Record_Of_Header.id)
              AND (pluc2a.ky_hieu IS NOT NULL);

            /*Thong tin phu luc 2B*/
            FOR vc_PLuc2B IN c_PLuc2B LOOP
            UPDATE qlt_gtrinh_gtgt_kt_02bc
               SET gia_tri_kkhai = vc_PLuc2B.gia_tri_ctieu
            WHERE (ctg_id = vc_PLuc2B.ctg_id);
            END LOOP;

            /*Thong tin phu luc 2C*/
            FOR vc_PLuc2C IN c_PLuc2C LOOP
            UPDATE qlt_gtrinh_gtgt_kt_02bc
               SET gia_tri_kkhai = vc_PLuc2C.gia_tri_ctieu
            WHERE (ctg_id = vc_PLuc2C.ctg_id);
            END LOOP;
            /*Ket thuc cap nhat thong tin cho to khai thay the*/
        ELSE
            /*Neu chua ton tai to khai trong ky ke khai*/
            IF ((p_Record_Of_Header.kykk_den_ngay + 1) = p_Record_Of_Header.kylb_tu_ngay) THEN
            /*Neu to khai dung han*/
                OPEN c_AnDinh_Exist; /*Kiem tra da co an dinh chua*/
                FETCH c_AnDinh_Exist INTO vc_AnDinh_Exist;
                IF (c_AnDinh_Exist%FOUND) THEN /*Neu co an dinh*/
                    v_Tthai_Tkhai := '3'; /*To khai nop cham sau an dinh*/
                ELSE /*Neu chua co an dinh*/
                    v_Tthai_Tkhai := '1'; /*To khai chinh thuc*/
                END IF;
                CLOSE c_AnDinh_Exist;
            ELSE
            /*Neu to khai khong dung han*/
                v_Tthai_Tkhai := '3'; /*To khai nop cham sau an dinh*/
            END IF;

            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Sinh giao dich*/
            Qlt_Pck_Gdich.Prc_Set_GTGT_2004;

            /*Thuc hien xu ly to khai*/
            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            /*Xu ly voi du lieu to khai Header*/
            INSERT INTO qlt_tkhai_hdr(id,
                                      ltd,
                                      tin,
                                      ten_dtnt,
                                      cqt_ma_cqt,
                                      hun_ma_tinh,
                                      hun_ma_huyen,
                                      dia_chi,
                                      ma_phong,
                                      ma_can_bo,
                                      dtk_ma_loai_tkhai,
                                      ngay_nop,
                                      kylb_tu_ngay,
                                      kylb_den_ngay,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      tthai,
                                      co_loi_ddanh,
                                      ghi_chu_loi,
                                      co_gtrinh_02a,
                                      co_gtrinh_02b,
                                      co_gtrinh_02c,
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   vc_Dtnt.ma_cqt,
                   vc_Dtnt.ma_tinh,
                   vc_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   vc_Dtnt.ma_phong,
                   vc_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   p_Record_Of_Header.co_gtrinh_02a,
                   p_Record_Of_Header.co_gtrinh_02b,
                   p_Record_Of_Header.co_gtrinh_02c,
                   p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat);

            /*Insert du lieu vao bang to khai Detail*/
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_gtgt_kt(id,
                                              tkh_id,
                                              tkh_ltd,
                                              ctk_id,
                                              so_tt,
                                              doanhso_dtnt,
                                              sothue_dtnt,
                                              doanhso_cqt,
                                              sothue_cqt,
                                              ke_khai_sai,
                                              ma_so_ct_thue,
                                              ma_so_ct_doanh_so)
                VALUES(v_Array_Of_Record_Dtl(i).id,
                       v_Hdr_Id,
                       0,
                       v_Array_Of_Record_Dtl(i).ctk_id,
                       v_Array_Of_Record_Dtl(i).so_tt,
                       v_Array_Of_Record_Dtl(i).doanhso_dtnt,
                       v_Array_Of_Record_Dtl(i).sothue_dtnt,
                       NVL(v_Array_Of_Record_Dtl(i).doanhso_cqt,0),
                       NVL(v_Array_Of_Record_Dtl(i).sothue_cqt,0),
                       v_Array_Of_Record_Dtl(i).ke_khai_sai,
                       v_Array_Of_Record_Dtl(i).ma_so_ct_thue,
                       v_Array_Of_Record_Dtl(i).ma_so_ct_doanhso);
            END LOOP;

            /*Dua so lieu phat sinh vao bang hach toan*/
            /***************************************************************/
            INSERT INTO qlt_htoan_tkhai_gtgt_kt_2004(id,
                                                     tkh_id,
                                                     tkh_ltd,
                                                     ccg_ma_cap,
                                                     ccg_ma_chuong,
                                                     lkn_ma_loai,
                                                     lkn_ma_khoan,
                                                     tmt_ma_muc,
                                                     tmt_ma_tmuc,
                                                     tmt_ma_thue,
                                                     thue_psinh_knay,
                                                     thue_ktru_knay,
                                                     thue_pnop_knay,
                                                     thue_ktru_ksau,
                                                     thue_dau_vao,
                                                     thue_dau_ra)
             VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    vc_Dtnt.cap,
                    vc_Dtnt.chuong,
                    vc_Dtnt.loai,
                    vc_Dtnt.khoan,
                    '014',
                    '01',
                    '01',
                    v_Thue_PSinh_KyNay,
                    v_Thue_KTru_KyNay,
                    v_Thue_PNop_KyNay,
                    v_Thue_KTru_KySau,
                    v_Thue_Dau_Vao,
                    v_Thue_Dau_Ra);
            /************************************************************************/
            /*Ket thuc dua so lieu phat sinh vao bang hach toan*/

            /*Thuc hien cap nhat du lieu cho cac phu luc giai trinh*/
            /*******************************************************/

            /*Xu ly du lieu ban giai trinh 2A*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02a(id,
                                               tkh_id,
                                               tkh_ltd,
                                               ctk_id,
                                               dien_giai,
                                               ky_hieu,
                                               kykk_tu_ngay,
                                               kykk_den_ngay,
                                               gia_tri_kkhai,
                                               gia_tri_dchinh,
                                               gia_tri_hhdv,
                                               gia_tri_thue,
                                               ghi_chu)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   v_Hdr_Id,
                   0,
                   DECODE(pluc2a.ky_hieu,'18',255,'19',255,'20',255,'21',255,
                                         '34',264,'35',264,'36',264,'37',264) ctk_id,
                   pluc2a.dien_giai,
                   pluc2a.ky_hieu,
                   TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH'),
                   LAST_DAY(TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH')),
                   pluc2a.gia_tri_slieu_kkhai,
                   pluc2a.gia_tri_slieu_dchinh,
                   pluc2a.gia_tri_hhdv,
                   pluc2a.gia_tri_thue_gtgt,
                   pluc2a.gia_tri_lydo_dchinh
            FROM rcv_v_tkhai_gtgt_kt_pluc2a pluc2a
            WHERE (pluc2a.hdr_id = p_Record_Of_Header.id)
              AND (pluc2a.ky_hieu IS NOT NULL);
            /*Ket thuc xu ly du lieu ban giai trinh 2A*/

            /*Xu ly du lieu ban giai trinh 2B*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02bc(id,
                                                tkh_id,
                                                tkh_ltd,
                                                ctg_id,
                                                gia_tri_kkhai)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   v_Hdr_Id,
                   0,
                   pluc2b.ctg_id,
                   pluc2b.gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2b pluc2b
            WHERE (pluc2b.hdr_id = p_Record_Of_Header.id);
            /*Ket thuc xu ly du lieu ban giai trinh 2B*/

            /*Xu ly du lieu ban giai trinh 2C*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02bc(id,
                                                tkh_id,
                                                tkh_ltd,
                                                ctg_id,
                                                gia_tri_kkhai)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   v_Hdr_Id,
                   0,
                   pluc2c.ctg_id,
                   pluc2c.gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2c pluc2c
            WHERE (pluc2c.hdr_id = p_Record_Of_Header.id);
            /*Ket thuc xu ly du lieu ban giai trinh 2C*/

            /******************************************************/
            /*Ket thuc cap nhat du lieu cho cac phu luc giai trinh*/
        END IF;
        /*Ket thuc xu ly to khai*/

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        COMMIT;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/11/2005
Muc dich: Thuc hien do du lieu to khai tu CSDL trung gian vao CSDL QLT
Tham so: NONE
*******************************************************************************/
    PROCEDURE Prc_Chuyen_Dlieu_Qlt IS
        /*Lay tat ca cac to khai hien co trong CSDL trung gian, chua duoc chuyen
         vao CSDL QLT (truong "da_nhan" co gia tri NULL)*/
        CURSOR c_Insert_Header_TKhai IS
            SELECT hdr.id,
                   hdr.tin,
                   hdr.ten_dtnt,
                   hdr.dia_chi,
                   tkhai.ma_tkhai_qlt loai_tkhai,
                   hdr.ngay_nop,
                   hdr.kylb_tu_ngay,
                   hdr.kylb_den_ngay,
                   hdr.kykk_tu_ngay,
                   hdr.kykk_den_ngay,
                   hdr.ngay_cap_nhat,
                   hdr.nguoi_cap_nhat,
                   DECODE(hdr.co_loi_ddanh,'x','Y',NULL) co_loi_ddanh,
                   hdr.so_hieu_tep,
                   hdr.so_tt_tk,
                   hdr.ghi_chu_loi,
                   hdr.co_gtrinh_02a,
                   hdr.co_gtrinh_02b,
                   hdr.co_gtrinh_02c
            FROM rcv_tkhai_hdr hdr,
                 rcv_map_tkhai tkhai
            WHERE (hdr.loai_tkhai = tkhai.ma_tkhai)
              AND (hdr.da_nhan IS NULL);
        vc_Insert_Header_TKhai c_Insert_Header_TKhai%ROWTYPE;
        vr_Record_Hdr Record_Hdr;
    BEGIN
        FOR vc_Insert_Header_TKhai IN c_Insert_Header_TKhai LOOP
            vr_Record_Hdr.id := vc_Insert_Header_TKhai.id;
            vr_Record_Hdr.tin := vc_Insert_Header_TKhai.tin;
            vr_Record_Hdr.ten_dtnt := vc_Insert_Header_TKhai.ten_dtnt;
            vr_Record_Hdr.dia_chi := vc_Insert_Header_TKhai.dia_chi;
            vr_Record_Hdr.loai_tkhai := vc_Insert_Header_TKhai.loai_tkhai;
            vr_Record_Hdr.ngay_nop := vc_Insert_Header_TKhai.ngay_nop;
            vr_Record_Hdr.kylb_tu_ngay := vc_Insert_Header_TKhai.kylb_tu_ngay;
            vr_Record_Hdr.kylb_den_ngay := vc_Insert_Header_TKhai.kylb_den_ngay;
            vr_Record_Hdr.kykk_tu_ngay := vc_Insert_Header_TKhai.kykk_tu_ngay;
            vr_Record_Hdr.kykk_den_ngay := vc_Insert_Header_TKhai.kykk_den_ngay;
            vr_Record_Hdr.ngay_cap_nhat := vc_Insert_Header_TKhai.ngay_cap_nhat;
            vr_Record_Hdr.nguoi_cap_nhat := vc_Insert_Header_TKhai.nguoi_cap_nhat;
            vr_Record_Hdr.co_loi_ddanh := vc_Insert_Header_TKhai.co_loi_ddanh;
            vr_Record_Hdr.so_hieu_tep := vc_Insert_Header_TKhai.so_hieu_tep;
            vr_Record_Hdr.so_tt_tk := vc_Insert_Header_TKhai.so_tt_tk;
            vr_Record_Hdr.ghi_chu_loi := vc_Insert_Header_TKhai.ghi_chu_loi;
            vr_Record_Hdr.co_gtrinh_02a := vc_Insert_Header_TKhai.co_gtrinh_02a;
            vr_Record_Hdr.co_gtrinh_02b := vc_Insert_Header_TKhai.co_gtrinh_02b;
            vr_Record_Hdr.co_gtrinh_02c := vc_Insert_Header_TKhai.co_gtrinh_02c;
            
            Prc_Insert_Header_Detail(vr_Record_Hdr);
        END LOOP;
    END;
END;
/


-- End of DDL Script for Package Body QLT_OWNER.RCV_PCK_CHUYEN_DLIEU_QLT


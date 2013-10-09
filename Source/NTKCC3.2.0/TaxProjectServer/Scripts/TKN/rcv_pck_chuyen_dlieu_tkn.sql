-- Start of DDL Script for Package Body TKN_OWNER.RCV_PCK_CHUYEN_DLIEU_TKN
-- Generated 23-Dec-2005 17:14:43 from TKN_OWNER@TKN_92

CREATE OR REPLACE 
PACKAGE rcv_pck_chuyen_dlieu_tkn IS
/*******************************************************************************
Phien ban: 1.0
Nguoi lap: Nguyen Ta Anh
Ngay lap: 15/11/2005
Muc dich: Do du lieu tu CSDL trung gian QLT sau khi quet to khai vao TKN
*******************************************************************************/
  TYPE Record_TKhai_Rcv IS RECORD (id NUMBER(10,0),
                                   id_tkhai_rcv_qlt NUMBER(10,0),
                                   ma_dtnt VARCHAR2(14),
                                   ky_kekhai VARCHAR2(7),
                                   loai_tkhai VARCHAR2(7),
                                   ky_lapbo VARCHAR2(7),
                                   br_dt NUMBER(20,2),
                                   brgtgt_dt NUMBER(20,2),
                                   brgtgt_thue NUMBER(20,2),
                                   br0_dt NUMBER(20,2),
                                   br5_dt NUMBER(20,2),
                                   br5_thue NUMBER(20,2),
                                   br10_dt NUMBER(20,2),
                                   br10_thue NUMBER(20,2),
                                   br20_dt NUMBER(20,2),
                                   br20_thue NUMBER(20,2),
                                   mv_dt NUMBER(20,2),
                                   mvgtgt_thue NUMBER(20,2),
                                   ktru_thue NUMBER(20,2),
                                   pnop_thue NUMBER(20,2),
                                   kytruoc_thue NUMBER(20,2),
                                   nopthieu_thue NUMBER(20,2),
                                   nopthua_thue NUMBER(20,2),
                                   dnop_thue NUMBER(20,2),
                                   htra_thue NUMBER(20,2),
                                   pnop_th_thue NUMBER(20,2),
                                   ngay_nop DATE,
                                   ngay_nhap DATE,
                                   nnhap VARCHAR2(20),
                                   tt_ghiso VARCHAR2(1),
                                   gchu VARCHAR2(50),
                                   chua_ktru_kytruoc NUMBER(20,2),
                                   mv_thue NUMBER(20,2),
                                   nhapkhau_dt NUMBER(20,2),
                                   nhapkhau_thue NUMBER(20,2),
                                   tscd_dt NUMBER(20,2),
                                   tscd_thue NUMBER(20,2),
                                   mvgtgt_dt NUMBER(20,2),
                                   dchinh_mv_tang_dt NUMBER(20,2),
                                   dchinh_mv_tang_thue NUMBER(20,2),
                                   dchinh_mv_giam_dt NUMBER(20,2),
                                   dchinh_mv_giam_thue NUMBER(20,2),
                                   tong_ktru_thue NUMBER(20,2),
                                   brkct_dt NUMBER(20,2),
                                   dchinh_br_tang_dt NUMBER(20,2),
                                   dchinh_br_tang_thue NUMBER(20,2),
                                   dchinh_br_giam_dt NUMBER(20,2),
                                   dchinh_br_giam_thue NUMBER(20,2),
                                   pnop_thue1 NUMBER(20,2),
                                   ktru_thue_luyke NUMBER(20,2),
                                   denghi_htra_thue NUMBER(20,2),
                                   tong_ktru_kysau NUMBER(20,2),
                                   dien_thoai VARCHAR2(20),
                                   fax VARCHAR2(20),
                                   email VARCHAR2(100),
                                   mv_trongnuoc_dt NUMBER(20,2),
                                   mv_trongnuoc_thue NUMBER(20,2),
                                   tongthue_mvgtgt NUMBER(20,2),
                                   tongdt_brgtgt NUMBER(20,2),
                                   tongthue_brgtgt NUMBER(20,2),
                                   br_thue NUMBER(20,2),
                                   tong_ktru_thue1 NUMBER(20,2),
                                   co_gtrinh_02a VARCHAR2(1),
                                   co_gtrinh_02b VARCHAR2(1),
                                   co_gtrinh_02c VARCHAR2(1),
                                   co_loi_ddanh VARCHAR2(1),
                                   so_hieu_tep VARCHAR2(10),
                                   so_hieu_tkhai VARCHAR2(10),
                                   khong_psinh VARCHAR2(1));
  TYPE Array_Of_DLieu_TKhai IS TABLE OF Record_TKhai_Rcv INDEX BY BINARY_INTEGER;

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

  TYPE Array_Of_Detail_Value IS TABLE OF NUMBER(20,2) INDEX BY BINARY_INTEGER;

/******************************************************************************/
  PROCEDURE Prc_Chuyen_Dlieu_Tkn;
/******************************************************************************/
  PROCEDURE Prc_DLieu_TKhai_Detail(p_Id NUMBER
                                 , vArray_Of_Detail_Value OUT Array_Of_Detail_Value
                                 , p_Khong_PSinh OUT VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_Insert_Dlieu_Bang_TGian(p_Record_Of_Header Record_Hdr
                                      , vArray_Of_Detail_Value Array_Of_Detail_Value
                                      , p_Khong_PSinh VARCHAR2);
/******************************************************************************/
  FUNCTION Fnc_Lay_Quy(p_Ngay_Dau DATE) RETURN NUMBER;
/******************************************************************************/
  FUNCTION Fnc_Sinh_SHieu_Tep(p_Loai_TKhai VARCHAR2,
                              p_Ky_Ke_Khai DATE) RETURN VARCHAR2;
/******************************************************************************/
  FUNCTION Fnc_Sinh_SHieu_TKhai(p_SHieu_Tep VARCHAR2) RETURN VARCHAR2;
/******************************************************************************/
  PROCEDURE Prc_Ghi_TKhai_Hdr(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                              p_SHieu_Tep VARCHAR2,
                              p_TThai_TKhai VARCHAR2,
                              p_Success OUT BOOLEAN,
                              p_Id OUT NUMBER);
/******************************************************************************/
  PROCEDURE Prc_Ghi_TKhai_Dtl(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                              p_Id NUMBER);
/******************************************************************************/
  PROCEDURE Prc_Lay_Id_CTieu(p_Loai_TKhai VARCHAR2,
                             p_KyKK VARCHAR2,
                             p_Id VARCHAR2,
                             p_Ma OUT VARCHAR2,
                             p_CTieu_Id OUT NUMBER,
                             p_So_TT OUT NUMBER);
/******************************************************************************/
  PROCEDURE Prc_Chuyen_Dlieu_TKhai;
/******************************************************************************/
  FUNCTION Fnc_TKhai_Exist(p_Tin VARCHAR2,
                           p_Loai_TKhai VARCHAR2,
                           p_Ky_Ke_Khai DATE,
                           p_TThai_TKhai OUT VARCHAR2,
                           p_Max_Ltd OUT NUMBER) RETURN NUMBER;
/******************************************************************************/
  PROCEDURE Prc_TinhPS_GhiLoi(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                              p_Mode VARCHAR2,
                              p_Id_New NUMBER DEFAULT NULL,
                              p_Id_Old NUMBER DEFAULT NULL);
/******************************************************************************/
  PROCEDURE Prc_Backup_DLieu(p_Id IN NUMBER,
                             p_Max_Ltd NUMBER);
/******************************************************************************/
  PROCEDURE Prc_Ghi_PhuLuc(p_Id NUMBER,
                           p_Id_Header NUMBER);
/******************************************************************************/
  PROCEDURE Prc_Capnhat_DLieu(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                              p_Id NUMBER,
                              p_TThai_TKhai VARCHAR2,
                              p_Id_TKhai_TGian NUMBER);
END;
/

-- Grants for Package
GRANT EXECUTE ON rcv_pck_chuyen_dlieu_tkn TO tkn
/

CREATE PUBLIC SYNONYM rcv_pck_chuyen_dlieu_tkn
  FOR rcv_pck_chuyen_dlieu_tkn
/

CREATE OR REPLACE 
PACKAGE BODY rcv_pck_chuyen_dlieu_tkn IS
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien lay du lieu to khai chi tiet tu bang RCV_TKHAI_DTL (ben CSDL
          QLT), dua vao mang
Tham so:
      - p_Id: Id cua to khai trong bang RCV_TKHAI_HDR
      - vArray_Of_Detail_Value: Mang chua du lieu to khai detail
      - p_Khong_PSinh: Co hoat dong mua ban phat sinh trong ky khong
*******************************************************************************/
    PROCEDURE Prc_DLieu_TKhai_Detail(p_Id NUMBER
                                   , vArray_Of_Detail_Value OUT Array_Of_Detail_Value
                                   , p_Khong_PSinh OUT VARCHAR2) IS
        CURSOR c_Detail_TKhai IS
            SELECT dtl.ky_hieu,
                   dtl.gia_tri
            FROM rcv_tkhai_dtl@qlt dtl
            WHERE (hdr_id = p_Id)
              AND (dtl.ky_hieu <> 1)
              AND (loai_dlieu = '0101')
            ORDER BY TO_NUMBER(ky_hieu);
        vc_Detail_TKhai c_Detail_TKhai%ROWTYPE;
        vcount NUMBER(10) := 1;

    BEGIN
        FOR vc_Detail_TKhai IN c_Detail_TKhai LOOP
            vcount := vcount + 1;
            vArray_Of_Detail_Value(vcount) := NVL(vc_Detail_TKhai.gia_tri,0);
        END LOOP;

        SELECT MAX(DECODE(dtl.gia_tri,'x', 'Y', NULL)) INTO p_Khong_PSinh
        FROM rcv_tkhai_dtl@qlt dtl
        WHERE (hdr_id = p_Id)
          AND (dtl.ky_hieu = 1);
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien DO du lieu vao bang RCV_TKHAI_GTGT_KTRU_TKTN
Tham so:
      - p_Record_Of_Header: Bien ban ghi chua thong tin du lieu header
      - vArray_Of_Detail_Value: Chua du lieu to khai detail
      - p_Khong_PSinh: Phat sinh hoat dong mua ban trong ky khong
*******************************************************************************/
    PROCEDURE Prc_Insert_Dlieu_Bang_TGian(p_Record_Of_Header Record_Hdr
                                        , vArray_Of_Detail_Value Array_Of_Detail_Value
                                        , p_Khong_PSinh VARCHAR2) IS
        CURSOR c_Dtnt IS
            SELECT dtnt.ten_dtnt,
                   dtnt.dien_thoai,
                   dtnt.fax,
                   dtnt.email
            FROM tkn_v_dtnt dtnt
            WHERE (dtnt.tin = p_Record_Of_Header.tin);
        vc_Dtnt c_Dtnt%ROWTYPE;
    BEGIN
        OPEN c_Dtnt;
        FETCH c_Dtnt INTO vc_Dtnt;
        IF (c_Dtnt%NOTFOUND) THEN
            CLOSE c_Dtnt;
            RETURN;
        END IF;
        CLOSE c_Dtnt;
        
        INSERT INTO rcv_tkhai_gtgt_ktru_tktn(id,
                                             id_tkhai_rcv_qlt,
                                             ma_dtnt,
                                             ky_kekhai,
                                             loai_tkhai,
                                             ky_lapbo,
                                             ngay_nop,
                                             ngay_nhap,
                                             co_gtrinh_02a,
                                             co_gtrinh_02b,
                                             co_gtrinh_02c,
                                             co_loi_ddanh,
                                             so_hieu_tep,
                                             so_hieu_tkhai,
                                             dien_thoai,
                                             fax,
                                             email,
                                             br_dt,
                                             brgtgt_dt,
                                             brgtgt_thue,
                                             br0_dt,
                                             br5_dt,
                                             br5_thue,
                                             br10_dt,
                                             br10_thue,
                                             mv_dt,
                                             chua_ktru_kytruoc,
                                             mv_thue,
                                             nhapkhau_dt,
                                             nhapkhau_thue,
                                             dchinh_mv_tang_dt,
                                             dchinh_mv_tang_thue,
                                             dchinh_mv_giam_dt,
                                             dchinh_mv_giam_thue,
                                             brkct_dt,
                                             dchinh_br_tang_dt,
                                             dchinh_br_tang_thue,
                                             dchinh_br_giam_dt,
                                             dchinh_br_giam_thue,
                                             pnop_thue1,
                                             ktru_thue_luyke,
                                             denghi_htra_thue,
                                             tong_ktru_kysau,
                                             mv_trongnuoc_dt,
                                             mv_trongnuoc_thue,
                                             tongthue_mvgtgt,
                                             tongdt_brgtgt,
                                             tongthue_brgtgt,
                                             br_thue,
                                             tong_ktru_thue1,
                                             khong_psinh)
                                             
            VALUES(rcv_seq_tkhai.NEXTVAL,
                  p_Record_Of_Header.id,
                  p_Record_Of_Header.tin,
                  TO_CHAR(p_Record_Of_Header.kykk_tu_ngay,'MM/RRRR'),
                  '1',
                  TO_CHAR(p_Record_Of_Header.kylb_tu_ngay,'MM/RRRR'),
                  p_Record_Of_Header.ngay_nop,
                  SYSDATE,
                  p_Record_Of_Header.co_gtrinh_02a,
                  p_Record_Of_Header.co_gtrinh_02b,
                  p_Record_Of_Header.co_gtrinh_02c,
                  p_Record_Of_Header.co_loi_ddanh,
                  p_Record_Of_Header.so_hieu_tep,
                  p_Record_Of_Header.so_tt_tk,
                  vc_Dtnt.dien_thoai,
                  vc_Dtnt.fax,
                  vc_Dtnt.email,
                  vArray_Of_Detail_Value(19),
                  vArray_Of_Detail_Value(22),
                  vArray_Of_Detail_Value(23),
                  vArray_Of_Detail_Value(24),
                  vArray_Of_Detail_Value(25),
                  vArray_Of_Detail_Value(26),
                  vArray_Of_Detail_Value(27),
                  vArray_Of_Detail_Value(28),
                  vArray_Of_Detail_Value(5),
                  vArray_Of_Detail_Value(2),
                  vArray_Of_Detail_Value(6),
                  vArray_Of_Detail_Value(9),
                  vArray_Of_Detail_Value(10),
                  vArray_Of_Detail_Value(12),
                  vArray_Of_Detail_Value(13),
                  vArray_Of_Detail_Value(14),
                  vArray_Of_Detail_Value(15),
                  vArray_Of_Detail_Value(21),
                  vArray_Of_Detail_Value(30),
                  vArray_Of_Detail_Value(31),
                  vArray_Of_Detail_Value(32),
                  vArray_Of_Detail_Value(33),
                  vArray_Of_Detail_Value(37),
                  vArray_Of_Detail_Value(38),
                  vArray_Of_Detail_Value(39),
                  vArray_Of_Detail_Value(40),
                  vArray_Of_Detail_Value(7),
                  vArray_Of_Detail_Value(8),
                  vArray_Of_Detail_Value(16),
                  vArray_Of_Detail_Value(34),
                  vArray_Of_Detail_Value(35),
                  vArray_Of_Detail_Value(20),
                  vArray_Of_Detail_Value(17),
                  p_Khong_PSinh);
                  
        -- Cap nhat trang thai bang trung gian QLT
/*
        UPDATE rcv_tkhai_hdr@qlt
        SET da_nhan = 'Y'
        WHERE (id = p_Record_Of_Header.id);

        COMMIT;
*/
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien tra quy may
Tham so:
      - p_Ngay_Dau: Ngay dau quy
*******************************************************************************/
	FUNCTION Fnc_Lay_Quy(p_Ngay_Dau DATE) RETURN NUMBER IS
		v_Quy NUMBER;
	BEGIN
		v_Quy := TO_NUMBER(TO_CHAR(p_Ngay_Dau, 'MM'));

		IF (v_Quy = 1) THEN
			RETURN 1;
		ELSIF (v_Quy = 4) THEN
			RETURN 2;
		ELSIF (v_Quy = 7) THEN
			RETURN 3;
		ELSIF (v_Quy = 10) THEN
			RETURN 4;
		ELSE
			RETURN NULL;
		END IF;
	END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien sinh ra so hieu tep TO khai
Tham so:
      - p_Loai_TKhai: Loai TO khai
      - p_Ky_Ke_Khai: Ky ke khai
*******************************************************************************/
    FUNCTION Fnc_Sinh_SHieu_Tep(p_Loai_TKhai VARCHAR2,
                                p_Ky_Ke_Khai DATE) RETURN VARCHAR2 IS
    	v_So_Hieu_Tep VARCHAR2(100);
    	
    	CURSOR c_Tep IS
    		SELECT	MAX(TO_NUMBER(SUBSTR(so_hieu,8)))
    		FROM tkn_tep_tkhai
    		WHERE (dtk_ma = p_Loai_TKhai)
    		  AND (so_hieu LIKE v_So_Hieu_Tep  || '%')
    		  AND (TO_CHAR(kykk_tu_ngay,'MM/RRRR') = TO_CHAR(p_Ky_Ke_Khai,'MM/RRRR'));
    	
    	CURSOR c_Loai_TKhai IS
    		SELECT loai_ky
    		FROM tkn_dm_tkhai
    		WHERE (ma = p_Loai_TKhai);
    	
    	v_Kieu_Ky VARCHAR2(1);
    	v_Count	NUMBER(10);
    BEGIN		
    	-- Lay kieu ky
    	OPEN c_Loai_TKhai;
    	FETCH c_Loai_TKhai INTO v_Kieu_Ky;
    	CLOSE	c_Loai_TKhai;
    	IF (v_Kieu_Ky = 'M') THEN
    		v_So_Hieu_Tep := TO_CHAR(p_Ky_Ke_Khai, 'RRMM') || p_Loai_TKhai;
    	ELSIF (v_Kieu_Ky = 'Q') THEN
    		v_So_Hieu_Tep := TO_CHAR(p_Ky_Ke_Khai, 'RR') || '0' ||
                             TO_CHAR(Fnc_Lay_Quy(p_Ky_Ke_Khai)) ||
                             p_Loai_TKhai;
    	END IF;
    	
    	-- Sinh so hieu tep theo ky ke khai
    	OPEN c_Tep;
    	FETCH c_Tep INTO v_Count;
    	IF (v_Count IS NOT NULL) THEN
    		v_So_Hieu_Tep := v_So_Hieu_Tep || '-' || TO_CHAR(v_Count + 1);
    	ELSE
    	v_So_Hieu_Tep := v_So_Hieu_Tep || '-1';
    	END IF;
    	CLOSE c_Tep;	
    	RETURN v_So_Hieu_Tep;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien sinh ra so hieu TO khai
Tham so:
      - p_SHieu_Tep: So hieu tep
*******************************************************************************/
    FUNCTION Fnc_Sinh_SHieu_TKhai(p_SHieu_Tep VARCHAR2) RETURN VARCHAR2 IS
    	CURSOR c_TKhai IS
    	   SELECT MAX(TO_NUMBER(so_hieu))
    	   FROM	tkn_tkhai_hdr
    	   WHERE (tep_so_hieu = p_SHieu_Tep);
    						
    	v_Num NUMBER;
    BEGIN
    	v_Num := NULL;
    	OPEN c_TKhai;
    	FETCH c_TKhai INTO v_Num;
    	CLOSE c_TKhai;
    	IF (v_Num IS NOT NULL) THEN
    		RETURN TO_CHAR(v_Num + 1);
    	ELSE
    		RETURN '1';
    	END IF;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 04/12/2005
Muc dich: Thuc hien lay ma chi tieu trong bang TKN_DM_CTIEU_TKHAI
Tham so:
      - p_Loai_TKhai: Loai TO khai
      - p_KyKK: Ky ke khai
      - p_Id: ID cua chi tieu TO khai
      - p_Ma: Ma chi tieu TO khai
      - p_CTieu_Id: ID tra ve chi tieu TO khai
      - p_So_TT: So thu tu cua chi tieu TO khai
*******************************************************************************/
    PROCEDURE Prc_Lay_Id_CTieu(p_Loai_TKhai VARCHAR2,
                               p_KyKK VARCHAR2,
                               p_Id VARCHAR2,
                               p_Ma OUT VARCHAR2,
                               p_CTieu_Id OUT NUMBER,
                               p_So_TT OUT NUMBER) IS
    	CURSOR	c_Chi_Tieu IS
    		SELECT id,
                   ma,
                   so_tt
    		FROM tkn_dm_ctieu_tkhai
    		WHERE (dtk_ma = p_Loai_TKhai)
              AND (id = p_Id)
    		  AND (ngay_bat_dau <= TRUNC(TO_DATE(p_KyKK,'MM/RRRR'),'MONTH'))
    		  AND (ngay_ket_thuc >= TRUNC(TO_DATE(p_KyKK,'MM/RRRR'),'MONTH')
                   OR ngay_ket_thuc IS NULL);
    	v_CTieu_Id	NUMBER(10);
    	
    BEGIN
      OPEN c_Chi_Tieu;
      FETCH c_Chi_Tieu INTO p_CTieu_Id, p_Ma, p_So_TT;
      CLOSE c_Chi_Tieu;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 04/12/2005
Muc dich: Thuc hien kiem tra ton tai TO khai
Tham so:
      - p_Tin: Ma so thue
      - p_Loai_TKhai: Loai TO khai
      - p_Ky_Ke_Khai: Ky ke khai
      - p_TThai_TKhai: Tra ve trang thai to khai
						+ '01': Chua ton tai to khai, chua bi an dinh
						+ '02': Chua ton tai to khai, da bi an dinh
						+ '03': Da ton tai to khai chinh thuc
						+ '04': Da ton tai to khai nop cham sau an dinh
						+ '00': Da ton tai to khai HDR, chua nhap DTL
*******************************************************************************/
    FUNCTION Fnc_TKhai_Exist(p_Tin VARCHAR2,
                             p_Loai_TKhai VARCHAR2,
                             p_Ky_Ke_Khai DATE,
                             p_TThai_TKhai OUT VARCHAR2,
                             p_Max_Ltd OUT NUMBER) RETURN NUMBER IS
        -- Kiem tra to khai ton tai
    	CURSOR c_TKhai_Exist IS
    		SELECT	MAX(hdr.ltd) ltd,
                    id
    		FROM tkn_tkhai_hdr hdr
    		WHERE (hdr.dtk_ma = p_Loai_TKhai)
			  AND (hdr.kykk_tu_ngay = TRUNC(p_Ky_Ke_Khai, 'MONTH'))
			  AND (hdr.tin = p_Tin)
			GROUP BY id;
			  
    	-- Kiem tra to khai an dinh
        CURSOR c_DTNT_Andinh IS
    		SELECT id
    		FROM tkn_an_dinh_hdr
    		WHERE (tin = p_Tin)
    		  AND (TO_CHAR(kykk_tu_ngay, 'MM/RRRR') = TO_CHAR(TRUNC(p_Ky_Ke_Khai, 'MONTH'), 'MM/RRRR'))
    		  AND (dtk_ma = p_Loai_TKhai);			
    	
        -- Lay thong tin to khai DTL
        CURSOR c_TKhai_Dtl(p_Id Number) IS
            SELECT COUNT(1)
            FROM tkn_tkhai_dtl
            WHERE (tkh_id = p_Id)
              AND (tkh_ltd = 0);
                            	
    	v_Hdr_Id NUMBER;
    	v_Hdr_Ltd NUMBER;
    	v_Return NUMBER;
        v_Dtl_Count	NUMBER;
        v_ADinh_Id NUMBER;    	
    BEGIN
    	OPEN c_TKhai_Exist;
    	FETCH c_TKhai_Exist INTO v_Hdr_Ltd, v_Hdr_Id;
    	CLOSE c_TKhai_Exist;

        OPEN c_DTNT_Andinh;
        FETCH c_DTNT_Andinh INTO v_ADinh_Id;
        CLOSE c_DTNT_Andinh;
        
    	OPEN c_TKhai_Dtl(v_Hdr_Id);
    	FETCH c_TKhai_Dtl INTO v_Dtl_Count;
    	CLOSE c_TKhai_Dtl;

                	
        IF (v_Hdr_Id IS NULL) THEN
            IF (v_ADinh_Id IS NULL) THEN
    		  p_TThai_TKhai := '01';
    		ELSE
    		  p_TThai_TKhai := '02';
    		END IF;
            p_Max_Ltd := v_Hdr_Ltd;    		
    		RETURN NULL;
    	ELSE
            IF (NOT (v_Dtl_Count > 0)) THEN
    		  p_TThai_TKhai := '00';
    		ELSE
                IF (v_ADinh_Id IS NULL) THEN
        		  p_TThai_TKhai := '03';
        		ELSE
        		  p_TThai_TKhai := '04';
        		END IF;
    		END IF;
    		p_Max_Ltd := v_Hdr_Ltd;
    		RETURN v_Hdr_Id;
    	END IF;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 03/12/2005
Muc dich: Thuc hien ghi du lieu vao bang TKN_TKHAI_HDR
Tham so:
      - p_Record_TKhai_Rcv: Bien chua du lieu 1 ban ghi  cua bang
                            RCV_TKHAI_GTGT_KTRU_TKTN
      - p_SHieu_Tep: So hieu tep TO khai
      - p_TThai_TKhai: Trang thai to khai
      - p_Success: Ghi header co thanh cong khong
      - p_Id: Id cua to khai header tra ve
*******************************************************************************/
    PROCEDURE Prc_Ghi_TKhai_Hdr(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                                p_SHieu_Tep VARCHAR2,
                                p_TThai_TKhai VARCHAR2,
                                p_Success OUT BOOLEAN,
                                p_Id OUT NUMBER) IS
        CURSOR c_Tai_Khoan IS
        	SELECT id
        	FROM tkn_dm_tai_khoan
        	WHERE (loai = '01')
        	  AND (mac_dinh = 'Y');
        vc_Tai_Khoan c_Tai_Khoan%ROWTYPE;
        
        CURSOR c_Dtnt IS
            SELECT *
            FROM tkn_v_dtnt
            WHERE (tin = p_Record_TKhai_Rcv.ma_dtnt);
        vc_Dtnt c_Dtnt%ROWTYPE;
        
        v_SHieu_TKhai VARCHAR2(10);
        v_Chung_Tu VARCHAR2(100);
        v_So_TKhoan	VARCHAR2(10);
        v_Han_Nop DATE;
        v_Temp VARCHAR2(10);
    BEGIN
        OPEN c_Dtnt;
        FETCH c_Dtnt INTO vc_Dtnt;
        /*Sinh ID cho to khai*/
    	SELECT tkn_seq_kkhai_hdr.NEXTVAL INTO p_Id FROM dual;
       	/*Sinh so hieu to khai*/
       	v_SHieu_TKhai := Fnc_Sinh_SHieu_TKhai(p_SHieu_Tep);
    	/*Sinh so chung tu*/
    	v_Chung_Tu := p_SHieu_Tep || '-' || v_SHieu_TKhai;
    	
    	OPEN c_Tai_Khoan;
    	FETCH c_Tai_Khoan INTO v_So_TKhoan;
    	CLOSE c_Tai_Khoan;
    	
    	v_Temp := TO_CHAR(ADD_MONTHS(TRUNC(TO_DATE(p_Record_TKhai_Rcv.ky_kekhai,'MM/RRRR'),'MONTH'),1),'MM/RRRR');
    	v_Temp := '25/' || v_Temp;
    	v_Han_Nop := TO_DATE(v_Temp, 'DD/MM/RRRR');
    	
        INSERT INTO tkn_tkhai_hdr(id
						 		 ,ltd
								 ,dtk_ma
								 ,tep_so_hieu
								 ,so_hieu
								 ,tkn_id
								 ,han_nop
								 ,so_ctu
								 ,tin
								 ,ten_dtnt
								 ,cqt_ma
								 ,cqt_ltd
								 ,ma_phong
								 ,ma_canbo
								 ,tih_ma
								 ,tih_ltd
								 ,hun_ma
								 ,hun_ltd
								 ,dia_chi
								 ,dien_thoai
								 ,fax
								 ,email
								 ,kykk_tu_ngay
								 ,kykk_den_ngay
							 	 ,kylb_tu_ngay
								 ,kylb_den_ngay
								 ,ngay_nop
								 ,ngay_nhap
								 ,khong_psinh
								 ,loi_so_hoc
					 			 ,loi_dinh_danh
								 ,ghi_chu
								 ,tthai
								 ,ngay_gdich
								 ,goc_dchinh
								 ,khong_ktra_dso_thang
								 ,co_dchinh    								
								 ,co_dchinh_02b
								 ,co_dchinh_02c)
        VALUES(p_Id
			,0
			,'01'
			,p_SHieu_Tep
			,v_SHieu_TKhai
			,v_So_TKhoan
			,v_Han_Nop
			,v_Chung_Tu
			,p_Record_TKhai_Rcv.ma_dtnt
			,vc_Dtnt.ten_dtnt
			,vc_Dtnt.cqt_ma
			,0
			,vc_Dtnt.ma_phong
			,vc_Dtnt.ma_canbo
			,vc_Dtnt.tih_ma
			,0
			,vc_Dtnt.hun_ma
			,0
			,vc_Dtnt.dia_chi
			,p_Record_TKhai_Rcv.dien_thoai
			,p_Record_TKhai_Rcv.fax
			,p_Record_TKhai_Rcv.email
			,TRUNC(TO_DATE(p_Record_TKhai_Rcv.ky_kekhai,'MM/RRRR'),'MONTH')
			,LAST_DAY(TRUNC(TO_DATE(p_Record_TKhai_Rcv.ky_kekhai,'MM/RRRR'),'MONTH'))
			,TRUNC(TO_DATE(p_Record_TKhai_Rcv.ky_lapbo, 'MM/RRRR'),'MONTH')
			,LAST_DAY(TRUNC(TO_DATE(p_Record_TKhai_Rcv.ky_lapbo, 'MM/RRRR'),'MONTH'))
			,p_Record_TKhai_Rcv.ngay_nop
			,p_Record_TKhai_Rcv.ngay_nhap
			,p_Record_TKhai_Rcv.khong_psinh
			,NULL
			,p_Record_TKhai_Rcv.co_loi_ddanh
			,NULL
			,p_TThai_TKhai
			,SYSDATE
			,NULL
			,NULL
			,p_Record_TKhai_Rcv.co_gtrinh_02a    			
			,p_Record_TKhai_Rcv.co_gtrinh_02b
			,p_Record_TKhai_Rcv.co_gtrinh_02c);
        CLOSE c_Dtnt;			
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 04/12/2005
Muc dich: Thuc hien ghi du lieu vao bang TKN_TKHAI_DTL
Tham so: 
      - p_Record_TKhai_Rcv: Bien chua du lieu 1 ban ghi  cua bang
                            RCV_TKHAI_GTGT_KTRU_TKTN
      - p_Id: Id cua to khai header
*******************************************************************************/
    PROCEDURE Prc_Ghi_TKhai_Dtl(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                                p_Id NUMBER) IS
    	v_CTieu_Id NUMBER(10);
    	v_So_TT	NUMBER(10);
    	v_Ma VARCHAR2(10);
    BEGIN
    	-- Chi tieu 11
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 45, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.chua_ktru_kytruoc,p_Record_TKhai_Rcv.chua_ktru_kytruoc,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 12
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 46, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL, p_Id, 0, v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.mv_dt,p_Record_TKhai_Rcv.mv_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 13
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 47, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.mv_thue,p_Record_TKhai_Rcv.mv_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 14
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 48, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.mv_trongnuoc_dt,p_Record_TKhai_Rcv.mv_trongnuoc_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 15
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 49, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.mv_trongnuoc_thue,p_Record_TKhai_Rcv.mv_trongnuoc_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 16
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 50, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.nhapkhau_dt,p_Record_TKhai_Rcv.nhapkhau_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 17
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 51, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.nhapkhau_thue,p_Record_TKhai_Rcv.nhapkhau_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 18
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 52, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_mv_tang_dt,p_Record_TKhai_Rcv.dchinh_mv_tang_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 19
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 53, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_mv_tang_thue,p_Record_TKhai_Rcv.dchinh_mv_tang_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 20
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 54, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_mv_giam_dt,p_Record_TKhai_Rcv.dchinh_mv_giam_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 21
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 55, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_mv_giam_thue,p_Record_TKhai_Rcv.dchinh_mv_giam_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 22
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 56, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.tongthue_mvgtgt,p_Record_TKhai_Rcv.tongthue_mvgtgt,NULL,'Y',SYSDATE);

    	-- Chi tieu 23
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 57, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.tong_ktru_thue1,p_Record_TKhai_Rcv.tong_ktru_thue1,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 24
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 58, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br_dt,p_Record_TKhai_Rcv.br_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 25
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 59, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br_thue,p_Record_TKhai_Rcv.br_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 26
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 60, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.brkct_dt,p_Record_TKhai_Rcv.brkct_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 27
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 61, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.brgtgt_dt,p_Record_TKhai_Rcv.brgtgt_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 28
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 62, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.brgtgt_thue,p_Record_TKhai_Rcv.brgtgt_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 29
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 63, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br0_dt,p_Record_TKhai_Rcv.br0_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 30
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 64, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br5_dt,p_Record_TKhai_Rcv.br5_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 31
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 65, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br5_thue,p_Record_TKhai_Rcv.br5_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 32
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 66, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br10_dt,p_Record_TKhai_Rcv.br10_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 33
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 67, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.br10_thue,p_Record_TKhai_Rcv.br10_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 34
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 68, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_br_tang_dt,p_Record_TKhai_Rcv.dchinh_br_tang_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 35
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 69, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_br_tang_thue,p_Record_TKhai_Rcv.dchinh_br_tang_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 36
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 70, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_br_giam_dt,p_Record_TKhai_Rcv.dchinh_br_giam_dt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 37
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 71, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.dchinh_br_giam_thue,p_Record_TKhai_Rcv.dchinh_br_giam_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 38
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 72, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.tongdt_brgtgt,p_Record_TKhai_Rcv.tongdt_brgtgt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 39
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 73, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.tongthue_brgtgt,p_Record_TKhai_Rcv.tongthue_brgtgt,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 40
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 74, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.pnop_thue1,p_Record_TKhai_Rcv.pnop_thue1,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 41
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 75, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.ktru_thue_luyke,p_Record_TKhai_Rcv.ktru_thue_luyke,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 42
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 76, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.denghi_htra_thue,p_Record_TKhai_Rcv.denghi_htra_thue,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu 43
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 77, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,p_Record_TKhai_Rcv.tong_ktru_kysau,p_Record_TKhai_Rcv.tong_ktru_kysau,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 78, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma, v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 79, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 80, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 81, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 82, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    	
    	-- Chi tieu *
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 83, v_Ma, v_CTieu_Id, v_So_TT);
    	INSERT INTO tkn_tkhai_dtl(id,tkh_id,tkh_ltd,ctk_id,ma,so_tt,gia_tri_kkhai,gia_tri_cqt,loi_so_hoc,co_thay_doi,ngay_gdich)
    	VALUES(tkn_seq_kkhai_dtl.NEXTVAL,p_Id,0,v_CTieu_Id,v_Ma,v_So_TT,NULL,NULL,NULL,'Y',SYSDATE);
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 07/12/2005
Muc dich: Thuc hien ghi cac phu luc
Tham so: 
      - p_Id: Id so hieu to khai trong bang trung gian
      - p_Id_Header: Id to tkhai header
*******************************************************************************/
    PROCEDURE Prc_Ghi_PhuLuc(p_Id NUMBER,
                             p_Id_Header NUMBER) IS
    BEGIN
        /*Phu luc 2A*/
        INSERT INTO tkn_gtrinh_dchinh(id,
                                      tkh_id,
                                      tkh_ltd,
                                      ctk_id,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      gia_tri_kkhai,
                                      gia_tri_dchinh,
                                      gia_tri_clech,
                                      so_thue_clech,
                                      lydo_dchinh)
        SELECT tkn_seq_kkhai_dtl.NEXTVAL,
               p_Id_Header,
               0,
               DECODE(pluc2a.ky_hieu,'18',52,'19',53,'20',54,'21',55,
                                     '34',68,'35',69,'36',70,'37',71) ctk_id,
               TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH'),
               LAST_DAY(TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH')),
               pluc2a.gia_tri_slieu_kkhai,
               pluc2a.gia_tri_slieu_dchinh,
               pluc2a.gia_tri_hhdv,
               pluc2a.gia_tri_thue_gtgt,
               pluc2a.gia_tri_lydo_dchinh
        FROM rcv_v_tkhai_gtgt_kt_pluc2a@qlt pluc2a
        WHERE (pluc2a.hdr_id = p_Id)
          AND (pluc2a.ky_hieu IS NOT NULL);

        /*Phu luc 2B*/
        INSERT INTO tkn_gtrinh_dchinh_02b(id,
                                          tkh_id,
                                          tkh_ltd,
                                          cgt_id,
                                          so_tt,
                                          gia_tri_kkhai)
        SELECT tkn_seq_gdich.NEXTVAL,
               p_Id_Header,
               0,
               pluc2b.ctg_id,
               pluc2b.so_tt - 3 so_tt,
               pluc2b.gia_tri_ctieu
        FROM rcv_v_tkhai_gtgt_kt_pluc2b@qlt pluc2b
        WHERE (pluc2b.hdr_id = p_Id);

        /*Phu luc 2C*/
        INSERT INTO tkn_gtrinh_dchinh_02c(id,
                                          tkh_id,
                                          tkh_ltd,
                                          cgt_id,
                                          so_tt,
                                          gia_tri_kkhai)
        SELECT tkn_seq_gdich.NEXTVAL,
               p_Id_Header,
               0,
               pluc2c.ctg_id,
               pluc2c.so_tt - 3 so_tt,
               pluc2c.gia_tri_ctieu
        FROM rcv_v_tkhai_gtgt_kt_pluc2c@qlt pluc2c
        WHERE (pluc2c.hdr_id = p_Id);
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 04/12/2005
Muc dich: Tinh so phat sinh va thuc hien ghi loi
Tham so: 
      - p_Record_TKhai_Rcv: Chua du lieu to khai trung gian
      - p_Mode: Mode Insert hay Update
      - p_Id_New: Id to khai khi insert moi
      - p_Id_Old: Id to khai khi update
*******************************************************************************/
    PROCEDURE Prc_TinhPS_GhiLoi(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                                p_Mode VARCHAR2,
                                p_Id_New NUMBER DEFAULT NULL,
                                p_Id_Old NUMBER DEFAULT NULL) IS
    	TYPE tab_TKhai IS TABLE OF NUMBER(20) INDEX BY BINARY_INTEGER;
    	TYPE tab_DSach_Loi IS TABLE OF VARCHAR2(10) INDEX BY BINARY_INTEGER;
    	
    	vtab_Tkhai tab_TKhai;
    	vtab_DSach_Loi tab_DSach_Loi;
    	
    	v_CTieu_Ma BINARY_INTEGER;
    	v_CTieu_Id NUMBER(10) := 0;
    	v_So_TT	NUMBER(10) := 0;
    	v_Ma NUMBER(10) := 0;
    	
    	v_Gia_Tri NUMBER(20,2) := 0;
    	v_SThue_KTru NUMBER(20,2) := 0;
    	v_Da_KTru_TKy NUMBER(20,2) := 0;
    	
    	v_Index NUMBER := 0;
    	v_Id_PSinh NUMBER := 0;
    	v_Loi_So_Hoc VARCHAR2(1);
    	v_Temp NUMBER := 0;
    	v_Id_Loi_TKhai NUMBER(10);
    	v_Id_Temp NUMBER(10);
    	v_Co_Loi Varchar2(1);
    	v_Nguong_Min NUMBER(10);
    	
        CURSOR c_Dtnt IS
            SELECT dtnt.ccg_ma_cap,
                   dtnt.ccg_ma_chuong,
                   dtnt.lkn_ma_loai,
                   dtnt.lkn_ma_khoan
            FROM tkn_v_dtnt dtnt
            WHERE (dtnt.tin = p_Record_TKhai_Rcv.ma_dtnt);
        vc_Dtnt c_Dtnt%ROWTYPE;
    	
    	-- Lay gia tri khau tru ky truoc chuyen sang
    	CURSOR c_So_KTru_Tiep IS
    		SELECT ktru_cky
    		FROM tkn_khau_tru
    		WHERE (tin = p_Record_TKhai_Rcv.ma_dtnt)
    		  AND (kykk_tu_ngay = TRUNC(ADD_MONTHS(TO_DATE(p_Record_TKhai_Rcv.ky_kekhai, 'MM/RRRR'), -1), 'MONTH'));

    	-- Lay gia chi ke khai cua chi tieu
    	CURSOR c_CTieu(vc_CTieu_Id NUMBER) IS
    		SELECT gia_tri_kkhai
    		FROM tkn_tkhai_dtl
    		WHERE (tkh_id = v_Id_Temp)
    		  AND (tkh_ltd = 0)
    		  AND (ctk_id = vc_CTieu_Id);
    BEGIN
        -----------------------------------------------
        /* Tao mang danh sach gia tri chi tieu to khai */
    	vtab_TKhai(11) := NVL(p_Record_TKhai_Rcv.chua_ktru_kytruoc, 0);
    	vtab_TKhai(12) := NVL(p_Record_TKhai_Rcv.mv_dt, 0);
    	vtab_TKhai(13) := NVL(p_Record_TKhai_Rcv.mv_thue, 0);
    	vtab_TKhai(14) := NVL(p_Record_TKhai_Rcv.mv_trongnuoc_dt, 0);
    	vtab_TKhai(15) := NVL(p_Record_TKhai_Rcv.mv_trongnuoc_thue, 0);
    	vtab_TKhai(16) := NVL(p_Record_TKhai_Rcv.nhapkhau_dt, 0);
    	vtab_TKhai(17) := NVL(p_Record_TKhai_Rcv.nhapkhau_thue, 0);
    	vtab_TKhai(18) := NVL(p_Record_TKhai_Rcv.dchinh_mv_tang_dt, 0);
    	vtab_TKhai(19) := NVL(p_Record_TKhai_Rcv.dchinh_mv_tang_thue, 0);
    	vtab_TKhai(20) := NVL(p_Record_TKhai_Rcv.dchinh_mv_giam_dt, 0);
    	vtab_TKhai(21) := NVL(p_Record_TKhai_Rcv.dchinh_mv_giam_thue, 0);
    	vtab_TKhai(22) := NVL(p_Record_TKhai_Rcv.tongthue_mvgtgt, 0);
    	vtab_TKhai(23) := NVL(p_Record_TKhai_Rcv.tong_ktru_thue1, 0);
    	vtab_TKhai(24) := NVL(p_Record_TKhai_Rcv.br_dt, 0);
    	vtab_TKhai(25) := NVL(p_Record_TKhai_Rcv.br_thue, 0);
    	vtab_TKhai(26) := NVL(p_Record_TKhai_Rcv.brkct_dt, 0);
    	vtab_TKhai(27) := NVL(p_Record_TKhai_Rcv.brgtgt_dt, 0);
    	vtab_TKhai(28) := NVL(p_Record_TKhai_Rcv.brgtgt_thue, 0);
    	vtab_TKhai(29) := NVL(p_Record_TKhai_Rcv.br0_dt, 0);
    	vtab_TKhai(30) := NVL(p_Record_TKhai_Rcv.br5_dt, 0);
    	vtab_TKhai(31) := NVL(p_Record_TKhai_Rcv.br5_thue, 0);
    	vtab_TKhai(32) := NVL(p_Record_TKhai_Rcv.br10_dt, 0);
    	vtab_TKhai(33) := NVL(p_Record_TKhai_Rcv.br10_thue, 0);
    	vtab_TKhai(34) := NVL(p_Record_TKhai_Rcv.dchinh_br_tang_dt, 0);
    	vtab_TKhai(35) := NVL(p_Record_TKhai_Rcv.dchinh_br_tang_thue, 0);
    	vtab_TKhai(36) := NVL(p_Record_TKhai_Rcv.dchinh_br_giam_dt, 0);
    	vtab_TKhai(37) := NVL(p_Record_TKhai_Rcv.dchinh_br_giam_thue, 0);
    	vtab_TKhai(38) := NVL(p_Record_TKhai_Rcv.tongdt_brgtgt, 0);
    	vtab_TKhai(39) := NVL(p_Record_TKhai_Rcv.tongthue_brgtgt, 0);
    	vtab_TKhai(40) := NVL(p_Record_TKhai_Rcv.pnop_thue1, 0);
    	vtab_TKhai(41) := NVL(p_Record_TKhai_Rcv.ktru_thue_luyke, 0);
    	vtab_TKhai(42) := NVL(p_Record_TKhai_Rcv.denghi_htra_thue, 0);
    	vtab_TKhai(43) := NVL(p_Record_TKhai_Rcv.tong_ktru_kysau, 0);

    	/* Tinh so phat sinh va ghi vao bang phat sinh */
    	OPEN c_Dtnt;
    	FETCH c_Dtnt INTO vc_Dtnt;
    	IF (c_Dtnt%NOTFOUND) THEN
    	   RETURN;
    	END IF;
    	CLOSE c_Dtnt;
    	
    	v_Da_KTru_TKy := vtab_TKhai(11) + vtab_TKhai(23) - vtab_TKhai(42) - vtab_TKhai(43);
    	IF (v_Da_KTru_TKy < 0) THEN
    		v_Da_KTru_TKy := 0;
    	END IF;
    	
        IF (p_Mode = 'I') THEN
            v_Id_Temp := p_Id_New;
    	SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Id_PSinh FROM dual;
    	INSERT INTO tkn_psinh_tkhai(id
    								,tkh_id
    								,tkh_ltd
    								,ccg_ma_cap
    								,ccg_ma_chuong
    								,ccg_ltd
    								,lkn_ma_loai
    								,lkn_ma_khoan
    								,lkn_ltd
    								,tmt_ma_muc
    								,tmt_ma_tmuc
    								,tmt_ma_thue
    								,tmt_ltd
    								,ktru_dky
    								,psinh_tky
    								,phai_nop_tky
    								,duoc_ktru_tky
    								,da_ktru_tky
    								,dnghi_hoan
    								,ktru_cky
    								,dso_dvao
    								,dso_dvao_thue
    								,thue_dvao
    								,dso_dra
    								,dso_dra_thue
    								,thue_dra
    								,ngay_gdich
    								,tyle_chiu_thue_tndn
    								,thue_suat
    								,miengiam_tndn)
        VALUES(v_Id_PSinh
    		  ,v_Id_Temp
    		  ,0
    		  ,vc_Dtnt.ccg_ma_cap
    		  ,vc_Dtnt.ccg_ma_chuong
    		  ,0
    		  ,vc_Dtnt.lkn_ma_loai
    		  ,vc_Dtnt.lkn_ma_khoan
    		  ,0
    		  ,'014'
    		  ,'01'
    		  ,'01'
    		  ,0
    		  ,vtab_TKhai(11)
    		  ,vtab_TKhai(39) -  vtab_TKhai(23)
    		  ,vtab_TKhai(40)
    		  ,vtab_TKhai(23) + vtab_TKhai(11)
    		  ,v_Da_KTru_TKy
    		  ,vtab_TKhai(42)
    		  ,vtab_TKhai(43)
    		  ,vtab_TKhai(12) + vtab_TKhai(18) + vtab_TKhai(20)
    		  ,vtab_TKhai(12) + vtab_TKhai(18) + vtab_TKhai(20)
    		  ,vtab_TKhai(23)
    		  ,vtab_TKhai(38)
    		  ,vtab_TKhai(38) - vtab_TKhai(26)
    		  ,vtab_TKhai(39)
    		  ,SYSDATE
    		  ,NULL
    		  ,NULL
    		  ,NULL);
        ELSIF (p_Mode = 'U') THEN
            v_Id_Temp := p_Id_Old;
            Update tkn_psinh_tkhai
            Set ktru_dky = vtab_TKhai(11)
				,psinh_tky = vtab_TKhai(39) -  vtab_TKhai(23)
				,phai_nop_tky = vtab_TKhai(40)
				,duoc_ktru_tky = vtab_TKhai(23) + vtab_TKhai(11)
				,da_ktru_tky = v_Da_KTru_TKy
				,dnghi_hoan = vtab_TKhai(42)
				,ktru_cky = vtab_TKhai(43)
				,dso_dvao = vtab_TKhai(12) + vtab_TKhai(18) + vtab_TKhai(20)
				,dso_dvao_thue = vtab_TKhai(12) + vtab_TKhai(18) + vtab_TKhai(20)
				,thue_dvao = vtab_TKhai(23)
				,dso_dra = vtab_TKhai(38)
				,dso_dra_thue = vtab_TKhai(38) - vtab_TKhai(26)
				,thue_dra = vtab_TKhai(39)
            Where TKH_ID = p_Id_Old
            And TKH_LTD = 0;
        END IF;
        
    	/* Kiem tra va ghi loi to khai */
    	OPEN c_So_KTru_Tiep;
    	FETCH c_So_KTru_Tiep INTO v_SThue_KTru;
    	CLOSE c_So_KTru_Tiep;
    	IF (v_SThue_KTru IS NULL) THEN
    		v_SThue_KTru := 0;
    	END IF;

    	--Chi tieu 11 = So thue GTGT con duoc khau tru ky truoc chuyen sang
       	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 45, v_Ma, v_CTieu_Id, v_So_TT);
        IF (vtab_TKhai(11) <> v_SThue_KTru) THEN
            v_Loi_So_Hoc := 'Y';
            vtab_DSach_Loi(v_Index) := '01';
            v_Index := v_Index + 1;
            v_Co_Loi := 'Y';
        ELSE
            v_Co_Loi := Null;
        END IF;

    	UPDATE tkn_tkhai_dtl
    	SET	gia_tri_cqt = v_SThue_KTru
    	   ,loi_so_hoc = v_Co_Loi
    	WHERE (tkh_id = v_Id_Temp)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);

/*    	
    	--Chi tieu 12 khong duoc khac 14 + 16
    	IF (vtab_Tkhai(12)) <> vtab_Tkhai(14) + vtab_Tkhai(16) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 46, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(14) + vtab_Tkhai(16)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);  		

    		v_Loi_So_Hoc := 'Y';
      	    vtab_DSach_Loi(v_Index) := '27';
      	    v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 13 khong duoc khac 15 + 17
    	IF (vtab_Tkhai(13) <> vtab_Tkhai(15) + vtab_Tkhai(17)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 47, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(15) + vtab_Tkhai(17)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '28';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 22 khong duoc khac 13 + 19 - 21
    	IF (vTab_Tkhai(22) <> vtab_Tkhai(13) + vtab_Tkhai(19) - vtab_Tkhai(21)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 56, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(13) + vtab_Tkhai(19) - vtab_Tkhai(21)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '29';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 24 khong duoc khac 26 + 27
    	IF (vtab_Tkhai(24) <> vtab_Tkhai(26) + vtab_Tkhai(27)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 58, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(26) + vtab_Tkhai(27)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '30';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 25 khong duoc khac 28
    	IF (vtab_Tkhai(25) <> vtab_Tkhai(28)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 59, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(28)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc :='Y';
    		vtab_Dsach_Loi(v_Index) :='31';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 27 khong duoc khac 29 + 30 +32
    	IF (vtab_Tkhai(27) <> vtab_Tkhai(29) + vtab_Tkhai(30) + vtab_Tkhai(32)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 61, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(29) + vtab_Tkhai(30) + vtab_Tkhai(32)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '32';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi Tieu 28 khong duoc khac 31 + 33
    	IF (vtab_Tkhai(28) <> vtab_Tkhai(31)+ vtab_Tkhai(33) ) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 62, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(31)+ vtab_Tkhai(33)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '33';
    		v_Index := v_Index + 1;
    	END IF;
*/    	
    	--Chi Tieu [30]*5% - Nguong <= |[31]| <= [30]*5% + Nguong
        v_Temp := ROUND((vtab_TKhai(30)*5)/100);
        v_Nguong_Min := v_Temp/10000;
        IF (v_Nguong_Min > 100000) THEN
            v_Nguong_Min := 100000;
        END IF;
        
       	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 65, v_Ma, v_CTieu_Id, v_So_TT);
        IF ((ABS(vtab_TKhai(31)) < v_Temp - v_Nguong_Min) OR (ABS(vtab_TKhai(31)) > v_Temp + v_Nguong_Min)) THEN
            v_Loi_So_Hoc := 'Y';
            vtab_DSach_Loi(v_Index) := '34';
            v_Index := v_Index + 1;
            v_Co_Loi := 'Y';
        ELSE
            v_Co_Loi := Null;
        END IF;
    	
    	UPDATE tkn_tkhai_dtl
    	SET	gia_tri_cqt = v_Temp
    	   ,loi_so_hoc = v_Co_Loi
    	WHERE (tkh_id = v_Id_Temp)
        AND (tkh_ltd = 0)
        AND (ctk_id = v_CTieu_Id);

    	--Chi Tieu [32]*10% - Nguong <= |[33]| <= [32]*10% + Nguong
    	v_Temp := ROUND((vtab_Tkhai(32)*10)/100);
        v_Nguong_Min := v_Temp/10000;
        IF (v_Nguong_Min > 100000) THEN
            v_Nguong_Min := 100000;
        END IF;
            	
   		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 67, v_Ma, v_CTieu_Id, v_So_TT);

    	IF ((ABS(vtab_TKhai(33)) < v_Temp - v_Nguong_Min) OR (ABS(vtab_TKhai(33)) > v_Temp + v_Nguong_Min)) THEN
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '35';
    		v_Index := v_Index + 1;
    		v_Co_Loi := 'Y';
        ELSE
            v_Co_Loi := Null;
    	END IF;
		UPDATE tkn_tkhai_dtl
		SET gia_tri_cqt = v_Temp
		   ,loi_so_hoc = v_Co_Loi
		WHERE (tkh_id = v_Id_Temp)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);

/*
    	-- Chi Tieu 38 khong duoc khac 24 + 34 - 36
    	IF vtab_Tkhai(38) <> vtab_Tkhai(24)+ vtab_Tkhai(34) - vtab_Tkhai(36) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 72, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET gia_tri_cqt = vtab_Tkhai(24) + vtab_Tkhai(34) - vtab_Tkhai(36)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '36';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	-- Chi tieu 39 khong duoc khac 25 + 35 - 37
    	IF vtab_Tkhai(39) <> vtab_Tkhai(25) + vtab_Tkhai(35) - vtab_Tkhai(37) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 73, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(25) + vtab_Tkhai(35) - vtab_Tkhai(37)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '37';
    		v_Index := v_Index + 1;
    	END IF;

    	--Chi tieu 40 khong duoc khac  39 - 23 -11
    	IF vtab_Tkhai(40) <> vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 74, v_Ma, v_CTieu_Id, v_So_TT);
    		IF (vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11))>=0 THEN
    			UPDATE tkn_tkhai_dtl
    			SET	gia_tri_cqt = vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11)
    			   ,loi_so_hoc = 'Y'
    			WHERE (tkh_id = v_Id_Temp)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_CTieu_Id);
    		ELSE
    			UPDATE tkn_tkhai_dtl
    			SET	gia_tri_cqt = 0
    			   ,loi_so_hoc = 'Y'
    			WHERE (tkh_id = v_Id_Temp)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_CTieu_Id);
    		END IF;
    		
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '38';
    		v_Index := v_Index + 1;
    	END IF;
    	
    	--Chi tieu 41 khong duoc khac  39 - 23 -11
    	IF vtab_Tkhai(41) <> ABS(vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11)) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 75, v_Ma, v_CTieu_Id, v_So_TT);
    		IF (vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11))< 0 THEN
    			UPDATE tkn_tkhai_dtl
    			SET	gia_tri_cqt = ABS(vtab_Tkhai(39) - vtab_Tkhai(23) - vtab_Tkhai(11))
    			   ,loi_so_hoc = 'Y'
    			WHERE (tkh_id = v_Id_Temp)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_CTieu_Id);
    		ELSE
    			UPDATE tkn_tkhai_dtl
    			SET gia_tri_cqt = 0
    			   ,loi_so_hoc = 'Y'
    			WHERE (tkh_id = v_Id_Temp)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_CTieu_Id);
    		END IF;
    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '39';
    		v_Index := v_Index + 1;
    	END IF;	

    	--Chi tieu 43 khong duoc khac 41 -42
    	IF vtab_Tkhai(43) <> vtab_Tkhai(41) - vtab_Tkhai(42) THEN
    		Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 77, v_Ma, v_CTieu_Id, v_So_TT);
    		UPDATE tkn_tkhai_dtl
    		SET	gia_tri_cqt = vtab_Tkhai(41) - vtab_Tkhai(42)
    		   ,loi_so_hoc = 'Y'
    		WHERE (tkh_id = v_Id_Temp)
              AND (tkh_ltd = 0)
              AND (ctk_id = v_CTieu_Id);

    		v_Loi_So_Hoc := 'Y';
    		vtab_Dsach_Loi(v_Index) := '40';
    		v_Index := v_Index + 1;
    	END IF;		

    	-----------------------------------------------
    	/* Cap nhat to khai Header */
    	UPDATE tkn_tkhai_hdr
    	SET	loi_so_hoc = v_Loi_So_Hoc
    	WHERE (id = v_Id_Temp)
          AND (ltd = 0);
    	
    	/* Ghi danh sach loi vao bang loi to khai */
    	IF (p_Mode = 'U') THEN
    	   DELETE FROM tkn_loi_tkhai
    	   WHERE (tkh_id = v_Id_Temp)
    	     AND (tkh_ltd = 0);
    	END IF;
    	
    	FOR v_I IN 0..v_Index - 1 LOOP
    	SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Id_Loi_TKhai FROM dual;
    		INSERT INTO tkn_loi_tkhai(id,
    								  tkh_id,
    								  tkh_ltd,
    								  loi_ma,
    								  ngay_psinh,
    								  ngay_sua,
    								  nguoi_sua,
    								  tthai,
    								  ghi_chu)
    		VALUES(v_Id_Loi_TKhai,
    			   v_Id_Temp,
    			   0,
    			   vtab_DSach_Loi(v_I),
    			   TRUNC(SYSDATE, 'MONTH'),
    			   TRUNC(SYSDATE, 'MONTH'),
    			   'CQT',
    			   '01',
    			   NULL);
    	END LOOP;
    	
    	-- Xoa cac mang trung gian
    	vtab_DSach_Loi.DELETE();
        vtab_TKhai.DELETE();
        
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 07/12/2005
Muc dich: Thuc hien backup du lieu to khai, cac phu luc, bang phat sinh,
          loi to khai
Tham so: 
      - p_Id: Id to khai
      - p_Max_Ltd: Lan thay doi lon nhat hien tai
*******************************************************************************/
    PROCEDURE Prc_Backup_DLieu(p_Id NUMBER,
                               p_Max_Ltd NUMBER) IS
    	v_Temp NUMBER;

    	-- Lay thong tin to khai header	
    	CURSOR c_TKhai_Hdr IS
    		SELECT *
    		FROM tkn_tkhai_hdr
    		WHERE (id = p_Id)
    		  AND (ltd = 0);
    	vc_TKhai_Hdr c_TKhai_Hdr%ROWTYPE;

    	-- Lay thong tin chi tiet to khai
    	CURSOR c_TKhai_Dtl IS
    		SELECT *
    		FROM tkn_tkhai_dtl
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);

    	-- Lay du lieu phat sinh to khai
    	CURSOR c_PSinh_TKhai IS
    		SELECT *
    		FROM tkn_psinh_tkhai
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);

    	-- Lay danh sach loi to khai
    	CURSOR c_Loi_TKhai IS
    		SELECT *
    		FROM tkn_loi_tkhai
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);

    	-- Lay danh sach giai trinh dieu chinh 02A
    	CURSOR c_GTrinh_DChinh IS
    		SELECT *
    		FROM tkn_gtrinh_dchinh
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);
    		  
    	-- Lay danh sach phu luc mau 02B
    	CURSOR c_Gtrinh_Dchinh_02B IS
    		SELECT  *
    		FROM tkn_gtrinh_dchinh_02b
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);
    		  
    	-- Lay danh sach phu luc mau 02C	
    	CURSOR c_Gtrinh_Dchinh_02C IS
    		SELECT  *
    		FROM tkn_gtrinh_dchinh_02c
    		WHERE (tkh_id = p_Id)
    		  AND (tkh_ltd = 0);	
    						
    BEGIN
    	-- Backup to khai header
    	OPEN c_TKhai_Hdr;
    	FETCH c_TKhai_Hdr INTO vc_TKhai_Hdr;
    	CLOSE c_TKhai_Hdr;

		INSERT INTO tkn_tkhai_hdr(id
								, ltd
								, dtk_ma
								, tep_so_hieu
								, so_hieu
								, so_ctu
								, tin
								, ten_dtnt
								, cqt_ma
								, cqt_ltd
								, ma_phong
								, ma_canbo
								, tih_ma
								, tih_ltd
								, hun_ma
								, hun_ltd
								, dia_chi
								, dien_thoai
								, fax
								, email
								, kykk_tu_ngay
								, kykk_den_ngay
								, kylb_tu_ngay
								, kylb_den_ngay
								, ngay_nop
								, ngay_nhap
								, khong_psinh
								, co_dchinh
								, loi_so_hoc
								, loi_dinh_danh
								, ghi_chu
								, tthai
								, ngay_gdich
								, tkn_id
								, han_nop
								, goc_dchinh
								, khong_ktra_dso_thang
								, co_dchinh_02b
								, co_dchinh_02c
								, co_dchinh_01b
								, ten_dan_dtu)
        VALUES(vc_TKhai_Hdr.ID
			, p_Max_Ltd + 1
			, vc_TKhai_Hdr.dtk_ma
			, vc_TKhai_Hdr.tep_so_hieu
			, vc_TKhai_Hdr.so_hieu
			, vc_TKhai_Hdr.so_ctu
			, vc_TKhai_Hdr.tin
			, vc_TKhai_Hdr.ten_dtnt
			, vc_TKhai_Hdr.cqt_ma
			, vc_TKhai_Hdr.cqt_ltd
			, vc_TKhai_Hdr.ma_phong
			, vc_TKhai_Hdr.ma_canbo
			, vc_TKhai_Hdr.tih_ma
			, vc_TKhai_Hdr.tih_ltd
			, vc_TKhai_Hdr.hun_ma
			, vc_TKhai_Hdr.hun_ltd
			, vc_TKhai_Hdr.dia_chi
			, vc_TKhai_Hdr.dien_thoai
			, vc_TKhai_Hdr.fax
			, vc_TKhai_Hdr.email
			, vc_TKhai_Hdr.kykk_tu_ngay
			, vc_TKhai_Hdr.kykk_den_ngay
			, vc_TKhai_Hdr.kylb_tu_ngay
			, vc_TKhai_Hdr.kylb_den_ngay
			, vc_TKhai_Hdr.ngay_nop
			, vc_TKhai_Hdr.ngay_nhap
			, vc_TKhai_Hdr.khong_psinh
			, vc_TKhai_Hdr.co_dchinh
			, vc_TKhai_Hdr.loi_so_hoc
			, vc_TKhai_Hdr.loi_dinh_danh
			, vc_TKhai_Hdr.ghi_chu
			, vc_TKhai_Hdr.tthai
			, vc_TKhai_Hdr.ngay_gdich
			, vc_TKhai_Hdr.tkn_id
			, vc_TKhai_Hdr.han_nop
			, vc_TKhai_Hdr.goc_dchinh
			, vc_TKHai_Hdr.khong_ktra_dso_thang
			, vc_TKhai_Hdr.co_dchinh_02b
			, vc_TKhai_Hdr.co_dchinh_02c
			, vc_TKhai_Hdr.co_dchinh_01b
			, vc_TKhai_Hdr.ten_dan_dtu);
   	
    	-- Backup to khai detail
    	FOR vc_TKhai_Dtl IN c_TKhai_Dtl LOOP
    		SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Temp FROM dual;
    		INSERT INTO	tkn_tkhai_dtl(id
    								, tkh_id
    								, tkh_ltd
    								, ctk_id
    								, gia_tri_kkhai
    								, gia_tri_cqt
    								, loi_so_hoc
    								, co_thay_doi
    								, ma
    								, so_tt
    								, ngay_gdich)
    		VALUES(v_Temp
    			, p_Id
    			, p_Max_Ltd + 1
    			, vc_TKhai_Dtl.ctk_id
    			, vc_TKhai_Dtl.gia_tri_kkhai
    			, vc_TKhai_Dtl.gia_tri_cqt
    			, vc_TKhai_Dtl.loi_so_hoc
    			, vc_TKhai_Dtl.co_thay_doi
    			, vc_TKhai_Dtl.ma
    			, vc_TKhai_Dtl.so_tt
    			, vc_TKhai_Dtl.ngay_gdich);
    	END LOOP;

        -- Backup bang giai trinh dieu chinh 02A
        FOR vc_GTrinh_DChinh IN c_GTrinh_DChinh LOOP
            SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Temp FROM dual;
        	INSERT INTO	tkn_gtrinh_dchinh(id
    									, tkh_id
    									, tkh_ltd
    									, ctk_id
    									, kykk_tu_ngay
    									, kykk_den_ngay
    									, gia_tri_kkhai
    									, gia_tri_dchinh
    									, gia_tri_clech
    									, so_thue_clech
    									, lydo_dchinh)
    		VALUES(v_Temp
    			, p_Id
    			, p_Max_Ltd + 1
    			, vc_GTrinh_DChinh.ctk_id
    			, vc_GTrinh_DChinh.kykk_tu_ngay
    			, vc_GTrinh_DChinh.kykk_den_ngay
    			, vc_GTrinh_DChinh.gia_tri_kkhai
    			, vc_GTrinh_DChinh.gia_tri_dchinh
    			, vc_GTrinh_DChinh.gia_tri_clech
    			, vc_GTrinh_DChinh.so_thue_clech
    			, vc_GTrinh_DChinh.lydo_dchinh);
        END LOOP;

      -- Backup bang phat sinh to khai
    	FOR vc_PSinh_TKhai IN c_PSinh_TKhai LOOP
    		SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Temp FROM dual;
    		INSERT INTO	tkn_psinh_tkhai(id
        								, tkh_id
        								, tkh_ltd
        								, ccg_ma_cap
        								, ccg_ma_chuong
        								, ccg_ltd
        								, lkn_ma_loai
        								, lkn_ma_khoan
        								, lkn_ltd
        								, tmt_ma_muc
        								, tmt_ma_tmuc
        								, tmt_ma_thue
        								, tmt_ltd
        								, ktru_dky
        								, psinh_tky
        								, phai_nop_tky
        								, duoc_ktru_tky
        								, da_ktru_tky
        								, dnghi_hoan
        								, ktru_cky
        								, dso_dvao
        								, dso_dvao_thue
        								, dso_dra
        								, dso_dra_thue
        								, thue_dra
        								, tyle_chiu_thue_tndn
        								, thue_suat
        								, miengiam_tndn)
        	VALUES(v_Temp
    			, p_Id
    			, p_Max_Ltd + 1
    			, vc_PSinh_TKhai.ccg_ma_cap
    			, vc_PSinh_TKhai.ccg_ma_chuong
    			, vc_PSinh_TKhai.ccg_ltd
    			, vc_PSinh_TKhai.lkn_ma_loai
    			, vc_PSinh_TKhai.lkn_ma_khoan
    			, vc_PSinh_TKhai.lkn_ltd
    			, vc_PSinh_TKhai.tmt_ma_muc
    			, vc_PSinh_TKhai.tmt_ma_tmuc
    			, vc_PSinh_TKhai.tmt_ma_thue
    			, vc_PSinh_TKhai.tmt_ltd
    			, vc_PSinh_TKhai.ktru_dky
    			, vc_PSinh_TKhai.psinh_tky
    			, vc_PSinh_TKhai.phai_nop_tky
    			, vc_PSinh_TKhai.duoc_ktru_tky
    			, vc_PSinh_TKhai.da_ktru_tky
    			, vc_PSinh_TKhai.dnghi_hoan
    			, vc_PSinh_TKhai.ktru_cky
    			, vc_PSinh_TKhai.dso_dvao
    			, vc_PSinh_TKhai.dso_dvao_thue
    			, vc_PSinh_TKhai.dso_dra
    			, vc_PSinh_TKhai.dso_dra_thue
    			, vc_PSinh_TKhai.thue_dra
    			, vc_PSinh_TKhai.tyle_chiu_thue_tndn
    			, vc_PSinh_TKhai.thue_suat
    			, vc_PSinh_TKhai.miengiam_tndn);
    	END LOOP;

    	--Back up bang TKN_GTRINH_DCHINH_02B
    	FOR v_Gtrinh_Dchinh_02B IN c_Gtrinh_Dchinh_02B LOOP
    		SELECT tkn_seq_kkhai_dtl.NEXTVAL INTO v_Temp FROM dual;
    		INSERT INTO tkn_gtrinh_dchinh_02b(id
                              				, tkh_id
                              				, tkh_ltd
                              				, cgt_id
                              				, so_tt
                              				, gia_tri_kkhai
                              				, gia_tri_cqt
                              				, loi_so_hoc)
    		VALUES(v_Temp
    			 , p_Id
    			 , p_Max_Ltd + 1
    			 , v_Gtrinh_Dchinh_02B.cgt_id
    			 , v_Gtrinh_Dchinh_02B.so_tt
    			 , v_Gtrinh_Dchinh_02B.gia_tri_kkhai
    			 , v_Gtrinh_Dchinh_02B.gia_tri_cqt
    			 , v_Gtrinh_Dchinh_02B.loi_so_hoc);
    	END LOOP;

    	--Back up bang TKN_GTRINH_DCHINH_02C.
    	FOR v_Gtrinh_Dchinh_02C IN c_Gtrinh_Dchinh_02C LOOP
    		SELECT TKN_SEQ_KKHAI_DTL.NEXTVAL INTO v_Temp FROM DUAL;
    		INSERT INTO tkn_gtrinh_dchinh_02c(id
                      				        , tkh_id
                            				, tkh_ltd
                              				, cgt_id
                              				, so_tt
                              				, gia_tri_kkhai
                              				, gia_tri_cqt
                              				, loi_so_hoc)
    		VALUES(v_Temp
    			 , p_Id
    			 , p_Max_Ltd + 1
    			 , v_Gtrinh_Dchinh_02C.cgt_id
    			 , v_Gtrinh_Dchinh_02C.so_tt
    			 , v_Gtrinh_Dchinh_02C.gia_tri_kkhai
    			 , v_Gtrinh_Dchinh_02C.gia_tri_cqt
    			 , v_Gtrinh_Dchinh_02C.loi_so_hoc);
    	END LOOP;
    	
    	-- Backup bang danh sach loi to khai
    	FOR vc_Loi_TKhai IN c_Loi_TKhai LOOP
    		UPDATE tkn_loi_tkhai
    		SET tkh_ltd = p_Max_Ltd + 1
    		WHERE (id = vc_Loi_TKhai.id);
    	END LOOP;

    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 07/12/2005
Muc dich: Thuc hien cap nhat du lieu cho to khai va cac phu luc theo trang thai
          moi
Tham so:
      - p_Record_TKhai_Rcv: Bien ban ghi chua du lieu to khai
      - p_Id: Id cua to khai da ton tai
      - p_TThai_TKhai: Trang thai moi cua to khai
      - p_Id_TKhai_TGian: Id cua to khai trong bang trung gian
*******************************************************************************/
    PROCEDURE Prc_Capnhat_DLieu(p_Record_TKhai_Rcv Record_TKhai_Rcv,
                                p_Id NUMBER,
                                p_TThai_TKhai VARCHAR2,
                                p_Id_TKhai_TGian NUMBER) IS
    	v_CTieu_Id NUMBER(10);
    	v_So_TT	NUMBER(10);
    	v_Ma VARCHAR2(10);
    	
        CURSOR c_PLuc2B IS
            SELECT ctg_id,
                   gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2b@qlt
            WHERE (hdr_id = p_Id_TKhai_TGian);

        CURSOR c_PLuc2C IS
            SELECT ctg_id,
                   gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2c@qlt
            WHERE (hdr_id = p_Id_TKhai_TGian);
    BEGIN
        /*Cap nhat du lieu to khai header*/
        UPDATE tkn_tkhai_hdr
        SET ngay_nop = p_Record_TKhai_Rcv.ngay_nop
		   ,ngay_nhap = p_Record_TKhai_Rcv.ngay_nhap
		   ,khong_psinh = p_Record_TKhai_Rcv.khong_psinh
		   ,loi_dinh_danh = p_Record_TKhai_Rcv.co_loi_ddanh
		   ,tthai = p_TThai_TKhai
		   ,ngay_gdich = SYSDATE
		   ,co_dchinh = p_Record_TKhai_Rcv.co_gtrinh_02a		
		   ,co_dchinh_02b = p_Record_TKhai_Rcv.co_gtrinh_02b
		   ,co_dchinh_02c = p_Record_TKhai_Rcv.co_gtrinh_02c
		WHERE (id = p_Id)
          AND (ltd = 0);
          
        /*Cap nhat du lieu to khai detail*/	
    	-- Chi tieu 11
 	    Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 45, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.chua_ktru_kytruoc,
            gia_tri_cqt = p_Record_TKhai_Rcv.chua_ktru_kytruoc
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 12 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 46, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.mv_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.mv_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 13 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 47, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.mv_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.mv_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 14 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 48, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.mv_trongnuoc_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.mv_trongnuoc_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 15 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 49, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.mv_trongnuoc_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.mv_trongnuoc_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 16 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 50, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.nhapkhau_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.nhapkhau_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 17 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 51, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.nhapkhau_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.nhapkhau_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 18 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 52, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_mv_tang_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_mv_tang_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 19 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 53, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_mv_tang_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_mv_tang_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 20 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 54, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_mv_giam_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_mv_giam_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 21 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 55, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_mv_giam_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_mv_giam_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 22 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 56, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.tongthue_mvgtgt,
            gia_tri_cqt = p_Record_TKhai_Rcv.tongthue_mvgtgt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);

    	-- Chi tieu 23 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 57, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.tong_ktru_thue1,
            gia_tri_cqt = p_Record_TKhai_Rcv.tong_ktru_thue1
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 24 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 58, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.br_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 25 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 59, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.br_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 26 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 60, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.brkct_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.brkct_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 27 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 61, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.brgtgt_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.brgtgt_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 28 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 62, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.brgtgt_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.brgtgt_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 29 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 63, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br0_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.br0_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 30 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 64, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br5_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.br5_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 31 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 65, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br5_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.br5_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 32 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 66, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br10_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.br10_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 33 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 67, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.br10_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.br10_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 34 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 68, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_br_tang_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_br_tang_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 35 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 69, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_br_tang_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_br_tang_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 36 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 70, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_br_giam_dt,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_br_giam_dt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 37 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 71, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.dchinh_br_giam_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.dchinh_br_giam_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 38 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 72, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.tongdt_brgtgt,
            gia_tri_cqt = p_Record_TKhai_Rcv.tongdt_brgtgt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 39 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 73, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.tongthue_brgtgt,
            gia_tri_cqt = p_Record_TKhai_Rcv.tongthue_brgtgt
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 40 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 74, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.pnop_thue1,
            gia_tri_cqt = p_Record_TKhai_Rcv.pnop_thue1
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 41 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 75, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.ktru_thue_luyke,
            gia_tri_cqt = p_Record_TKhai_Rcv.ktru_thue_luyke
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 42 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 76, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.denghi_htra_thue,
            gia_tri_cqt = p_Record_TKhai_Rcv.denghi_htra_thue
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
    	-- Chi tieu 43 
    	Prc_Lay_Id_CTieu('01', p_Record_TKhai_Rcv.ky_kekhai, 77, v_Ma, v_CTieu_Id, v_So_TT);
    	UPDATE tkn_tkhai_dtl
        SET gia_tri_kkhai = p_Record_TKhai_Rcv.tong_ktru_kysau,
            gia_tri_cqt = p_Record_TKhai_Rcv.tong_ktru_kysau
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0)
          AND (ctk_id = v_CTieu_Id);
    	
        /*Cap nhat phu luc 02A*/
        /*Xoa ban ghi cu*/	
        DELETE FROM tkn_gtrinh_dchinh
        WHERE (tkh_id = p_Id)
          AND (tkh_ltd = 0);
        /*Insert ban ghi moi*/
        INSERT INTO tkn_gtrinh_dchinh(id,
                                      tkh_id,
                                      tkh_ltd,
                                      ctk_id,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      gia_tri_kkhai,
                                      gia_tri_dchinh,
                                      gia_tri_clech,
                                      so_thue_clech,
                                      lydo_dchinh)
        SELECT tkn_seq_kkhai_dtl.NEXTVAL,
               p_Id,
               0,
               DECODE(pluc2a.ky_hieu,'18',52,'19',53,'20',54,'21',55,
                                     '34',68,'35',69,'36',70,'37',71) ctk_id,
               TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH'),
               LAST_DAY(TRUNC(TO_DATE(pluc2a.gia_tri_ky_kkhai,'MM/RRRR'),'MONTH')),
               pluc2a.gia_tri_slieu_kkhai,
               pluc2a.gia_tri_slieu_dchinh,
               pluc2a.gia_tri_hhdv,
               pluc2a.gia_tri_thue_gtgt,
               pluc2a.gia_tri_lydo_dchinh
        FROM rcv_v_tkhai_gtgt_kt_pluc2a@qlt pluc2a
        WHERE (pluc2a.hdr_id = p_Id_TKhai_TGian)
          AND (pluc2a.ky_hieu IS NOT NULL);
        	
        /*Cap nhat phu luc 02B*/		
        FOR vc_PLuc2B IN c_PLuc2B LOOP
        UPDATE tkn_gtrinh_dchinh_02b
           SET gia_tri_kkhai = vc_PLuc2B.gia_tri_ctieu
        WHERE (cgt_id = vc_PLuc2B.ctg_id);
        END LOOP;

        /*Cap nhat phu luc 02C*/		
        FOR vc_PLuc2C IN c_PLuc2C LOOP
        UPDATE tkn_gtrinh_dchinh_02c
           SET gia_tri_kkhai = vc_PLuc2C.gia_tri_ctieu
        WHERE (cgt_id = vc_PLuc2C.ctg_id);
        END LOOP;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 02/12/2005
Muc dich: Thuc hien lay du lieu to khai tu bang RCV_TKHAI_HDR va VIEW
          RCV_V_TKHAI_GTGT_KT, do vao bang RCV_TKHAI_GTGT_KTRU_TKTN
Tham so: NONE
*******************************************************************************/
    PROCEDURE Prc_Chuyen_Dlieu_Tkn IS
        CURSOR c_Header_TKhai IS
            SELECT HDR.*
            FROM rcv_tkhai_hdr@qlt hdr
               , tkn_v_dtnt DTNT
            WHERE (hdr.da_nhan IS NULL)
            And DTNT.TIN = HDR.TIN
            ORDER BY hdr.id;
            
        vc_Header_TKhai c_Header_TKhai%ROWTYPE;
        vr_Record_Hdr Record_Hdr;
        vArray_Of_Detail_Value Array_Of_Detail_Value;
        v_Khong_PSinh VARCHAR2(1);
    BEGIN
        FOR vc_Header_TKhai IN c_Header_TKhai LOOP
            vr_Record_Hdr.id := vc_Header_TKhai.id;
            vr_Record_Hdr.tin := vc_Header_TKhai.tin;
            vr_Record_Hdr.ngay_nop := vc_Header_TKhai.ngay_nop;
            vr_Record_Hdr.kykk_tu_ngay := vc_Header_TKhai.kykk_tu_ngay;
            vr_Record_Hdr.kylb_tu_ngay := vc_Header_TKhai.kylb_tu_ngay;
            vr_Record_Hdr.co_gtrinh_02a := vc_Header_TKhai.co_gtrinh_02a;
            vr_Record_Hdr.co_gtrinh_02b := vc_Header_TKhai.co_gtrinh_02b;
            vr_Record_Hdr.co_gtrinh_02c := vc_Header_TKhai.co_gtrinh_02c;
            vr_Record_Hdr.so_hieu_tep := vc_Header_TKhai.so_hieu_tep;
            vr_Record_Hdr.so_tt_tk := vc_Header_TKhai.so_tt_tk;
            IF (vc_Header_TKhai.co_loi_ddanh = 'x') THEN
                vr_Record_Hdr.co_loi_ddanh := 'Y';
            ELSE
                vr_Record_Hdr.co_loi_ddanh := NULL;
            END IF;
            Prc_DLieu_TKhai_Detail(vc_Header_TKhai.id,
                                   vArray_Of_Detail_Value,
                                   v_Khong_PSinh);
                                   
            Prc_Insert_Dlieu_Bang_TGian(vr_Record_Hdr,
                                        vArray_Of_Detail_Value,
                                        v_Khong_PSinh);
        END LOOP;
        /*Chuyen du lieu to khai vao cac bang that*/
        Prc_Chuyen_Dlieu_TKhai;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 04/12/2005
Muc dich: Thuc hien chuyen du lieu TO khai tu bang RCV_TKHAI_GTGT_KTRU_TKTN vao
          cac bang TKN_TKHAI_HDR, TKN_TKHAI_DTL, TKN_GTRINH_DCHINH,
          TKN_GTRINH_DCHINH_02B, TKN_GTRINH_DCHINH_02C
Tham so: NONE
*******************************************************************************/
    PROCEDURE Prc_Chuyen_Dlieu_TKhai IS
        v_SHieu_Tep VARCHAR2(10);
        vr_Record_TKhai_Rcv Record_TKhai_Rcv;
        
        CURSOR c_DLieu_TKhai_TGian IS
            SELECT *
            FROM rcv_tkhai_gtgt_ktru_tktn rcv_tktn
            WHERE (rcv_tktn.da_nhan IS NULL)
            ORDER BY rcv_tktn.id;

        v_Array_Of_DLieu_TKhai Array_Of_DLieu_TKhai;
        v_Count NUMBER(10) := 0;
        v_Id_TKhai_Exist NUMBER;
        v_So_TKhai NUMBER(10) := 0; --So to khai da nhap
        v_Success BOOLEAN;
        v_TThai_TKhai VARCHAR2(10);
        v_Max_Ltd NUMBER(10);
        v_Id NUMBER(10);
    BEGIN
       	v_SHieu_Tep := Fnc_Sinh_SHieu_Tep('01',ADD_MONTHS(TRUNC(SYSDATE, 'MONTH'), -1));

        INSERT INTO tkn_tep_tkhai(so_hieu,
                                  dtk_ma,
                                  kykk_tu_ngay,
                                  kykk_den_ngay,
								  ngay_tao,
  								  so_tkhai,
								  so_tkhai_dnhap,
								  tthai)
		VALUES(v_SHieu_Tep,
			   '01',
			   ADD_MONTHS(TRUNC(SYSDATE, 'MONTH'), -1),
			   LAST_DAY(ADD_MONTHS(TRUNC(SYSDATE, 'MONTH'), -1)),
			   TRUNC(SYSDATE, 'MONTH'),
			   5000,
			   0,
			   'Y');

        FOR vc_DLieu_TKhai_TGian IN c_DLieu_TKhai_TGian LOOP
            v_Count := v_Count + 1;
            v_Array_Of_DLieu_TKhai(v_Count).id := vc_DLieu_TKhai_TGian.id;
            v_Array_Of_DLieu_TKhai(v_Count).id_tkhai_rcv_qlt := vc_DLieu_TKhai_TGian.id_tkhai_rcv_qlt;
            v_Array_Of_DLieu_TKhai(v_Count).ma_dtnt := vc_DLieu_TKhai_TGian.ma_dtnt;
            v_Array_Of_DLieu_TKhai(v_Count).ky_kekhai := vc_DLieu_TKhai_TGian.ky_kekhai;
            v_Array_Of_DLieu_TKhai(v_Count).loai_tkhai := vc_DLieu_TKhai_TGian.loai_tkhai;
            v_Array_Of_DLieu_TKhai(v_Count).ky_lapbo := vc_DLieu_TKhai_TGian.ky_lapbo;
            v_Array_Of_DLieu_TKhai(v_Count).br_dt := vc_DLieu_TKhai_TGian.br_dt;
            v_Array_Of_DLieu_TKhai(v_Count).brgtgt_dt := vc_DLieu_TKhai_TGian.brgtgt_dt;
            v_Array_Of_DLieu_TKhai(v_Count).brgtgt_thue := vc_DLieu_TKhai_TGian.brgtgt_thue;
            v_Array_Of_DLieu_TKhai(v_Count).br0_dt := vc_DLieu_TKhai_TGian.br0_dt;
            v_Array_Of_DLieu_TKhai(v_Count).br5_dt := vc_DLieu_TKhai_TGian.br5_dt;
            v_Array_Of_DLieu_TKhai(v_Count).br5_thue := vc_DLieu_TKhai_TGian.br5_thue;
            v_Array_Of_DLieu_TKhai(v_Count).br10_dt := vc_DLieu_TKhai_TGian.br10_dt;
            v_Array_Of_DLieu_TKhai(v_Count).br10_thue := vc_DLieu_TKhai_TGian.br10_thue;
            v_Array_Of_DLieu_TKhai(v_Count).br20_dt := vc_DLieu_TKhai_TGian.br20_dt;
            v_Array_Of_DLieu_TKhai(v_Count).br20_thue := vc_DLieu_TKhai_TGian.br20_thue;
            v_Array_Of_DLieu_TKhai(v_Count).mv_dt := vc_DLieu_TKhai_TGian.mv_dt;
            v_Array_Of_DLieu_TKhai(v_Count).mvgtgt_thue := vc_DLieu_TKhai_TGian.mvgtgt_thue;
            v_Array_Of_DLieu_TKhai(v_Count).ktru_thue := vc_DLieu_TKhai_TGian.ktru_thue;
            v_Array_Of_DLieu_TKhai(v_Count).pnop_thue := vc_DLieu_TKhai_TGian.pnop_thue;
            v_Array_Of_DLieu_TKhai(v_Count).kytruoc_thue := vc_DLieu_TKhai_TGian.kytruoc_thue;
            v_Array_Of_DLieu_TKhai(v_Count).nopthieu_thue := vc_DLieu_TKhai_TGian.nopthieu_thue;
            v_Array_Of_DLieu_TKhai(v_Count).nopthua_thue := vc_DLieu_TKhai_TGian.nopthua_thue;
            v_Array_Of_DLieu_TKhai(v_Count).dnop_thue := vc_DLieu_TKhai_TGian.dnop_thue;
            v_Array_Of_DLieu_TKhai(v_Count).htra_thue := vc_DLieu_TKhai_TGian.htra_thue;
            v_Array_Of_DLieu_TKhai(v_Count).pnop_th_thue := vc_DLieu_TKhai_TGian.pnop_th_thue;
            v_Array_Of_DLieu_TKhai(v_Count).ngay_nop := vc_DLieu_TKhai_TGian.ngay_nop;
            v_Array_Of_DLieu_TKhai(v_Count).ngay_nhap := vc_DLieu_TKhai_TGian.ngay_nhap;
            v_Array_Of_DLieu_TKhai(v_Count).nnhap := vc_DLieu_TKhai_TGian.nnhap;
            v_Array_Of_DLieu_TKhai(v_Count).tt_ghiso := vc_DLieu_TKhai_TGian.tt_ghiso;
            v_Array_Of_DLieu_TKhai(v_Count).gchu := vc_DLieu_TKhai_TGian.gchu;
            v_Array_Of_DLieu_TKhai(v_Count).chua_ktru_kytruoc := vc_DLieu_TKhai_TGian.chua_ktru_kytruoc;
            v_Array_Of_DLieu_TKhai(v_Count).mv_thue := vc_DLieu_TKhai_TGian.mv_thue;
            v_Array_Of_DLieu_TKhai(v_Count).nhapkhau_dt := vc_DLieu_TKhai_TGian.nhapkhau_dt;
            v_Array_Of_DLieu_TKhai(v_Count).nhapkhau_thue := vc_DLieu_TKhai_TGian.nhapkhau_thue;
            v_Array_Of_DLieu_TKhai(v_Count).tscd_dt := vc_DLieu_TKhai_TGian.tscd_dt;
            v_Array_Of_DLieu_TKhai(v_Count).tscd_thue := vc_DLieu_TKhai_TGian.tscd_thue;
            v_Array_Of_DLieu_TKhai(v_Count).mvgtgt_dt := vc_DLieu_TKhai_TGian.mvgtgt_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_mv_tang_dt := vc_DLieu_TKhai_TGian.dchinh_mv_tang_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_mv_tang_thue := vc_DLieu_TKhai_TGian.dchinh_mv_tang_thue;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_mv_giam_dt := vc_DLieu_TKhai_TGian.dchinh_mv_giam_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_mv_giam_thue := vc_DLieu_TKhai_TGian.dchinh_mv_giam_thue;
            v_Array_Of_DLieu_TKhai(v_Count).tong_ktru_thue := vc_DLieu_TKhai_TGian.tong_ktru_thue;
            v_Array_Of_DLieu_TKhai(v_Count).brkct_dt := vc_DLieu_TKhai_TGian.brkct_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_br_tang_dt := vc_DLieu_TKhai_TGian.dchinh_br_tang_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_br_tang_thue := vc_DLieu_TKhai_TGian.dchinh_br_tang_thue;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_br_giam_dt := vc_DLieu_TKhai_TGian.dchinh_br_giam_dt;
            v_Array_Of_DLieu_TKhai(v_Count).dchinh_br_giam_thue := vc_DLieu_TKhai_TGian.dchinh_br_giam_thue;
            v_Array_Of_DLieu_TKhai(v_Count).pnop_thue1 := vc_DLieu_TKhai_TGian.pnop_thue1;
            v_Array_Of_DLieu_TKhai(v_Count).ktru_thue_luyke := vc_DLieu_TKhai_TGian.ktru_thue_luyke;
            v_Array_Of_DLieu_TKhai(v_Count).denghi_htra_thue := vc_DLieu_TKhai_TGian.denghi_htra_thue;
            v_Array_Of_DLieu_TKhai(v_Count).tong_ktru_kysau := vc_DLieu_TKhai_TGian.tong_ktru_kysau;
            v_Array_Of_DLieu_TKhai(v_Count).dien_thoai := vc_DLieu_TKhai_TGian.dien_thoai;
            v_Array_Of_DLieu_TKhai(v_Count).fax := vc_DLieu_TKhai_TGian.fax;
            v_Array_Of_DLieu_TKhai(v_Count).email := vc_DLieu_TKhai_TGian.email;
            v_Array_Of_DLieu_TKhai(v_Count).mv_trongnuoc_dt := vc_DLieu_TKhai_TGian.mv_trongnuoc_dt;
            v_Array_Of_DLieu_TKhai(v_Count).mv_trongnuoc_thue := vc_DLieu_TKhai_TGian.mv_trongnuoc_thue;
            v_Array_Of_DLieu_TKhai(v_Count).tongthue_mvgtgt := vc_DLieu_TKhai_TGian.tongthue_mvgtgt;
            v_Array_Of_DLieu_TKhai(v_Count).tongdt_brgtgt := vc_DLieu_TKhai_TGian.tongdt_brgtgt;
            v_Array_Of_DLieu_TKhai(v_Count).tongthue_brgtgt := vc_DLieu_TKhai_TGian.tongthue_brgtgt;
            v_Array_Of_DLieu_TKhai(v_Count).br_thue := vc_DLieu_TKhai_TGian.br_thue;
            v_Array_Of_DLieu_TKhai(v_Count).tong_ktru_thue1 := vc_DLieu_TKhai_TGian.tong_ktru_thue1;
            v_Array_Of_DLieu_TKhai(v_Count).co_gtrinh_02a := vc_DLieu_TKhai_TGian.co_gtrinh_02a;
            v_Array_Of_DLieu_TKhai(v_Count).co_gtrinh_02b := vc_DLieu_TKhai_TGian.co_gtrinh_02b;
            v_Array_Of_DLieu_TKhai(v_Count).co_gtrinh_02c := vc_DLieu_TKhai_TGian.co_gtrinh_02c;
            v_Array_Of_DLieu_TKhai(v_Count).co_loi_ddanh := vc_DLieu_TKhai_TGian.co_loi_ddanh;
            v_Array_Of_DLieu_TKhai(v_Count).khong_psinh := vc_DLieu_TKhai_TGian.khong_psinh;

            -- Kiem tra ton tai to khai
            v_Id_TKhai_Exist := Fnc_TKhai_Exist(vc_DLieu_TKhai_TGian.ma_dtnt,
                                                '01',
                                                TO_DATE(vc_DLieu_TKhai_TGian.ky_kekhai,'MM/RRRR'),
                                                v_TThai_TKhai,
                                                v_Max_Ltd);
            -- Neu chua ton tai -> to khai chinh thuc
            IF (v_Id_TKhai_Exist IS NULL) THEN
                v_So_TKhai := v_So_TKhai + 1;
                -- Ghi to khai header
                Prc_Ghi_TKhai_Hdr(v_Array_Of_DLieu_TKhai(v_Count),
                                  v_SHieu_Tep,
                                  v_TThai_TKhai,
                                  v_Success,
                                  v_Id);
                -- Ghi to khai detail
                Prc_Ghi_TKhai_Dtl(v_Array_Of_DLieu_TKhai(v_Count), v_Id);
                -- Tinh phat sinh va ghi loi
                Prc_TinhPS_GhiLoi(v_Array_Of_DLieu_TKhai(v_Count), 'I', v_Id, Null);
                --Ghi cac phu luc
                Prc_Ghi_PhuLuc(vc_DLieu_TKhai_TGian.id_tkhai_rcv_qlt,v_Id);

            -- Neu da ton tai -> to khai thay the
            ELSE
                v_So_TKhai := v_So_TKhai + 1;
                -- Backup to khai va cac phu luc
                Prc_Backup_DLieu(v_Id_TKhai_Exist,v_Max_Ltd);
                -- Cap nhat to khai va cac phu luc theo trang thai moi
                Prc_Capnhat_DLieu(v_Array_Of_DLieu_TKhai(v_Count),
                                  v_Id_TKhai_Exist,
                                  v_TThai_TKhai,
                                  vc_DLieu_TKhai_TGian.id_tkhai_rcv_qlt);
                -- Tinh phat sinh va ghi loi
                Prc_TinhPS_GhiLoi(v_Array_Of_DLieu_TKhai(v_Count),'U',NULL,v_Id_TKhai_Exist);
            END IF;

            -- Cap nhat trang thai bang trung gian TKN
            UPDATE rcv_tkhai_gtgt_ktru_tktn
            SET da_nhan = 'Y'
            WHERE (id = vc_DLieu_TKhai_TGian.id);

            UPDATE rcv_tkhai_hdr@qlt
            SET da_nhan = 'Y'
            WHERE (id = vc_DLieu_TKhai_TGian.id_tkhai_rcv_qlt);
            
        END LOOP;

        IF (v_So_TKhai <> 0) THEN
        	UPDATE tkn_tep_tkhai
        	SET	so_tkhai = v_So_TKhai,
                so_tkhai_dnhap = v_So_TKhai,
                tthai = 'N'
        	WHERE (so_hieu = v_SHieu_Tep);
        ELSE
        	DELETE FROM tkn_tep_tkhai
        	WHERE (so_hieu = v_SHieu_Tep);
        END IF;

        COMMIT;
    END;
END;
/


-- End of DDL Script for Package Body TKN_OWNER.RCV_PCK_CHUYEN_DLIEU_TKN


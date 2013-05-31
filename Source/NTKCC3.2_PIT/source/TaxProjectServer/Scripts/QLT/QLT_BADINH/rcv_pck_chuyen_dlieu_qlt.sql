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
                            loai VARCHAR2(2),
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
                            co_gtrinh_02c CHAR(1),
                            phong_xly VARCHAR2(7));
/******************************************************************************/
  TYPE Record_Dtnt IS RECORD(tin VARCHAR2(14),
                             ten_dtnt VARCHAR2(60),
                             dia_chi VARCHAR2(60),
                             ma_cqt VARCHAR2(5),
                             ma_tinh VARCHAR2(3),
                             ma_huyen VARCHAR2(5),
                             ma_phong VARCHAR2(7),
                             ma_canbo VARCHAR2(10),
                             ma_cap VARCHAR2(1),
                             ma_chuong VARCHAR2(3),
                             ma_loai VARCHAR2(2),
                             ma_khoan VARCHAR2(2),
                             trang_thai VARCHAR2(2),
                             dien_thoai VARCHAR2(20),
                             fax VARCHAR2(20),
                             email VARCHAR2(30),
                             ngay_kdoanh DATE,
                             ngay_tchinh DATE,
                             loai VARCHAR2(2));
/******************************************************************************/
  PROCEDURE Prc_Chuyen_Dlieu_Qlt;
/******************************************************************************/
  FUNCTION Fnc_TKhai_Exits(p_Tin VARCHAR2,
                           p_Loai_TKhai VARCHAR2,
                           p_Kykk_Tu_Ngay DATE,
                           p_Kykk_Den_Ngay DATE,
                           p_Tthai OUT VARCHAR2,
                           p_Loai VARCHAR2) RETURN NUMBER;
/******************************************************************************/
  FUNCTION Fnc_Lay_So_Ktru_Ktruoc(p_Tin VARCHAR2,
                                  p_Kylb DATE) RETURN NUMBER;
/******************************************************************************/
  PROCEDURE Prc_TKhai_GTGT(p_Record_Of_Header Record_Hdr,
                           p_Record_Dtnt Record_Dtnt,
                           p_TKhai_Exits_Id NUMBER,
                           p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_TKhai_TNDN_Quy(p_Record_Of_Header Record_Hdr,
                               p_Record_Dtnt Record_Dtnt,
                               p_TKhai_Exits_Id NUMBER,
                               p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_TKhai_TNguyen(p_Record_Of_Header Record_Hdr,
                              p_Record_Dtnt Record_Dtnt,
                              p_TKhai_Exits_Id NUMBER,
                              p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_TKhai_Qtoan_TNguyen(p_Record_Of_Header Record_Hdr,
                                    p_Record_Dtnt Record_Dtnt,
                                    p_TKhai_Exits_Id NUMBER,
                                    p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_TKhai_TTDB(p_Record_Of_Header Record_Hdr,
                           p_Record_Dtnt Record_Dtnt,
                           p_TKhai_Exits_Id NUMBER,
                           p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  PROCEDURE Prc_Thong_Tin_Dtnt(p_Tin VARCHAR2,
                               p_Record_Dtnt OUT Record_Dtnt);
/******************************************************************************/
  FUNCTION Fnc_AnDinh_Exits(p_Tin VARCHAR2,
                            p_Loai_TKhai VARCHAR2,
                            p_Kykk_Tu_Ngay DATE,
                            p_Kykk_Den_Ngay DATE) RETURN BOOLEAN;
/******************************************************************************/
  PROCEDURE Prc_TKhai_QToan_TNDN_Nam(p_Record_Of_Header Record_Hdr,
                                     p_Record_Dtnt Record_Dtnt,
                                     p_TKhai_Exits_Id NUMBER,
                                     p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  FUNCTION Fnc_So_KKhai_TKhai_Quy (p_Tin VARCHAR2,
    						       p_Tmt_Ma_Muc VARCHAR2,
        						   p_Tmt_Ma_TMuc VARCHAR2,
        						   p_Tmt_Ma_Thue VARCHAR2,
        						   p_KyKK DATE) RETURN NUMBER;
/******************************************************************************/
  FUNCTION Fnc_So_KKhai_Tnguyen(p_Tin IN VARCHAR2,
                                p_Tmt_Ma_Muc IN VARCHAR2,
                                p_Tmt_Ma_TMuc IN VARCHAR2,
                                p_Tmt_Ma_Thue IN VARCHAR2,
                                p_KyKK IN DATE ) RETURN qlt_xltk_gdich.so_tien%TYPE;
/******************************************************************************/
  FUNCTION Fnc_Han_Nop(p_Ma_TKhai VARCHAR2,
                       p_Loai VARCHAR2,
                       p_Cuoi_Ky DATE,
                       p_Nam_TChinh DATE DEFAULT NULL) RETURN DATE;
/******************************************************************************/
  PROCEDURE Prc_TKhai_GTGT_TT(p_Record_Of_Header Record_Hdr,
                              p_Record_Dtnt Record_Dtnt,
                              p_TKhai_Exits_Id NUMBER,
                              p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  FUNCTION Fnc_Ktra_Ctieu(p_So_DTNT NUMBER,
                          p_So_CQT  NUMBER,
                          p_KyLB DATE) RETURN BOOLEAN;
/******************************************************************************/
  PROCEDURE Prc_TKhai_QToan_GTGT_TT(p_Record_Of_Header Record_Hdr,
                                    p_Record_Dtnt Record_Dtnt,
                                    p_TKhai_Exits_Id NUMBER,
                                    p_Tthai_Tkhai VARCHAR2);
/******************************************************************************/
  FUNCTION Fnc_So_KKhai_GTGT_TT(p_Tin VARCHAR2,
						  	    p_Tmt_Ma_Muc VARCHAR2,
						        p_Tmt_Ma_TMuc VARCHAR2,
					  		    p_Tmt_Ma_Thue VARCHAR2,
							    p_KyKK DATE ) RETURN NUMBER;
/******************************************************************************/
  PROCEDURE Prc_So_Nhan_Hoso(p_Record_Of_Header Record_Hdr,
                             p_Record_Dtnt Record_Dtnt,
                             p_Loai_Tkhai VARCHAR2,
                             p_Nhom VARCHAR2,
                             p_Ma_Pluc VARCHAR2);
/******************************************************************************/
  FUNCTION Fnc_Get_So_Hoso(p_Date IN DATE,
                           p_Ma IN VARCHAR2) RETURN VARCHAR2;
END;
/

CREATE OR REPLACE 
PACKAGE BODY rcv_pck_chuyen_dlieu_qlt IS
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/11/2005
Muc dich: Thuc hien kiem tra to khai hay quyet toan da ton tai trong QLT chua,
          ung voi mot ky ke khai
Tham so:
     - p_Tin: Ma so thue cua DTNT
     - p_Loai_TKhai: Ma to khai
     - p_Kykk_Tu_Ngay: Ky ke khai tu ngay
     - p_Kykk_Den_Ngay: Ky ke khai den ngay
     - p_Tthai: Trang thai tra ve
     - p_Loai: To khai hay quyet toan
*******************************************************************************/
    FUNCTION Fnc_TKhai_Exits(p_Tin VARCHAR2,
                             p_Loai_TKhai VARCHAR2,
                             p_Kykk_Tu_Ngay DATE,
                             p_Kykk_Den_Ngay DATE,
                             p_Tthai OUT VARCHAR2,
                             p_Loai VARCHAR2) RETURN NUMBER IS
        CURSOR c_TKhai_Exits IS
            SELECT  hdr.id
                   ,hdr.tthai
            FROM qlt_tkhai_hdr hdr
            WHERE (hdr.tin = p_Tin)
              AND (hdr.dtk_ma_loai_tkhai = p_Loai_TKhai)
              AND (hdr.kykk_tu_ngay = p_Kykk_Tu_Ngay)
              AND (hdr.kykk_den_ngay = p_Kykk_Den_Ngay)
              AND ((hdr.ltd = 0) OR (hdr.ltd = 1))
            ORDER BY HDR.LTD;

        CURSOR c_QToan_Exits IS
            SELECT hdr.id, hdr.tthai
            FROM qlt_qtoan_hdr hdr
            WHERE (hdr.tin = p_Tin)
              AND (hdr.dqt_ma = p_Loai_TKhai)
              AND (hdr.kykk_tu_ngay = p_Kykk_Tu_Ngay)
              AND (hdr.kykk_den_ngay = p_Kykk_Den_Ngay)
              AND (hdr.ltd = 0);

        vc_TKhai_Exits c_TKhai_Exits%ROWTYPE;
        vc_QToan_Exits c_QToan_Exits%ROWTYPE;
        v_ID    Number(10):= Null;
    BEGIN
        --Xu ly voi to khai
        IF (p_Loai = 'TK') THEN
            FOR vc_TKhai_Exits IN c_TKhai_Exits LOOP
                p_TThai := vc_TKhai_Exits.TTHAI;
                v_ID := vc_TKhai_Exits.ID;
            END LOOP;
            RETURN v_ID;

        --Xu ly voi quyet toan
        ELSIF (p_Loai = 'QT') THEN
            OPEN c_QToan_Exits;
            FETCH c_QToan_Exits INTO vc_QToan_Exits;
            IF (c_QToan_Exits%FOUND) THEN
                p_Tthai := vc_QToan_Exits.tthai;
                RETURN vc_QToan_Exits.id;
            ELSE
                RETURN NULL;
            END IF;
            CLOSE c_QToan_Exits;
        END IF;
    END;
/******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/01/2006
Muc dich: Thuc hien kiem tra an dinh da ton tai trong QLT chua,
          ung voi mot ky ke khai
Tham so:
     - p_Tin: Ma so thue cua DTNT
     - p_Loai_TKhai: Loai to khai cua DTNT nop
     - p_KyKK: Ky ke khai ma DTNT nop to khai
*******************************************************************************/
    FUNCTION Fnc_AnDinh_Exits(p_Tin VARCHAR2,
                              p_Loai_TKhai VARCHAR2,
                              p_Kykk_Tu_Ngay DATE,
                              p_Kykk_Den_Ngay DATE) RETURN BOOLEAN IS
        CURSOR c_AnDinh_Exist IS
            SELECT 1
            FROM qlt_ds_an_dinh_hdr hdr
            WHERE (hdr.tin = p_Tin)
              AND (hdr.dtk_ma = p_Loai_TKhai)
              AND (hdr.kykk_tu_ngay = p_Kykk_Tu_Ngay)
              AND (hdr.kykk_den_ngay = p_Kykk_Den_Ngay);
        v_Found_AnDinh BOOLEAN := FALSE;
        test number;
    BEGIN
        OPEN c_AnDinh_Exist; /*Kiem tra da co an dinh chua*/
        FETCH c_AnDinh_Exist INTO test;
        IF (c_AnDinh_Exist%FOUND) THEN
            v_Found_AnDinh := TRUE;
        END IF;
        CLOSE c_AnDinh_Exist;
        RETURN v_Found_AnDinh;
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
Ngay lap: 19/01/2006
Muc dich: Thuc hien lay so ke khai tam nop quy cua DTNT trong 12 thang
Tham so:
        - p_Tin: Ma so thue cua DTNT
        - p_Tmt_Ma_Muc: Ma muc
        - p_Tmt_Ma_TMuc: Ma tieu muc
        - p_Tmt_Ma_Thue: Ma thue
        - p_KyKK: Ky ke khai
*******************************************************************************/
    FUNCTION Fnc_So_KKhai_TKhai_Quy (p_Tin VARCHAR2,
    							     p_Tmt_Ma_Muc VARCHAR2,
    							     p_Tmt_Ma_TMuc VARCHAR2,
    							     p_Tmt_Ma_Thue VARCHAR2,
    							     p_KyKK DATE) RETURN NUMBER IS
      --Lay tong so thue tam nop tren to khai TNDN nam
    	CURSOR c_Thue_Tam_Nop IS
    		SELECT SUM(DECODE(dgd_ma_gdich, 58, (-1) * NVL(so_tien,0), NVL(so_tien,0)))
    		FROM qlt_xltk_gdich
    		WHERE (tin = p_Tin)
    		  AND (tmt_ma_muc = p_Tmt_Ma_Muc)
    		  AND (tmt_ma_tmuc = p_Tmt_Ma_TMuc)
    		  AND (tmt_ma_thue = p_Tmt_Ma_Thue)
    		  AND (dgd_kieu_gdich IN ('21','25')) --Kieu giao dich va
    		  AND (dgd_ma_gdich IN ('04','16','20','78','22','24','58')) -- ma giao dich
    		  AND (kykk_tu_ngay >= TRUNC(p_KyKK, 'RRRR'))
    		  AND (kykk_den_ngay <= TO_DATE('31/12/' || TO_CHAR(p_KyKK,'RRRR'),'DD/MM/RRRR'));

        --Lay tong so thue tam nop tren to khai TNDN thang
        CURSOR c_Thue_Tam_Nop_Thang IS
        	SELECT SUM(DECODE(dgd_ma_gdich, 58, (-1) * NVL(so_tien,0), NVL(so_tien,0)))
        	FROM qlt_xltk_gdich
        	WHERE (tin = p_Tin)
        	  AND (tmt_ma_muc = p_Tmt_Ma_Muc)
        	  AND (tmt_ma_tmuc = p_Tmt_Ma_TMuc)
        	  AND (tmt_ma_thue = p_Tmt_Ma_Thue)
        	  AND (dgd_kieu_gdich IN ('21','25')) --Kieu giao dich va
        	  AND (dgd_ma_gdich IN ('06','12','18','22','24','58')) -- ma giao dich
        	  AND (kykk_tu_ngay >= TRUNC(p_KyKK, 'RRRR'))
        	  AND (kykk_den_ngay <= TO_DATE('31/12/' || TO_CHAR(p_KyKK,'RRRR'),'DD/MM/RRRR'));

    	vc_Thue_Tam_Nop	NUMBER;
    	vc_Thue_TBao NUMBER;
    	vc_Thue_Tam_Nop_Thang NUMBER;
    	v_So_KKhai_TKhai NUMBER;
    	v_Exist_TKhai_DChinh BOOLEAN;
    BEGIN
        --Lay so thue tam nop ca nam tren to khai TNDN nam
        OPEN c_Thue_Tam_Nop; --So thue tam nop tren to khai dieu chinh
        FETCH c_Thue_Tam_Nop INTO vc_Thue_Tam_Nop;
        CLOSE c_Thue_Tam_Nop;

        OPEN c_Thue_Tam_Nop_Thang; --So thue tam nop tren to khai TNDN thang
        FETCH c_Thue_Tam_Nop_Thang INTO vc_Thue_Tam_Nop_Thang;
        CLOSE c_Thue_Tam_Nop_Thang;

        --Tinh so ke khai tren to khai tam nop
        v_So_KKhai_TKhai := NVL(vc_Thue_Tam_Nop,0) + NVL(vc_Thue_Tam_Nop_Thang,0);
        RETURN(v_So_KKhai_TKhai);
    END;
/*******************************************************************************
Nguoi lap: Khainhg
Ngay lap: 19/04/2006
Noi dung: Lay so ke khai 12 thang tren to khai thue Tai Nguyen
Tham so:
        - p_Tin: Ma so thue cua DTNT
        - p_Tmt_Ma_Muc: Ma muc
        - p_Tmt_Ma_TMuc: Ma tieu muc
        - p_Tmt_Ma_Thue: Ma thue
        - p_KyKK: Ky ke khai

********************************************************************************/
FUNCTION Fnc_So_KKhai_Tnguyen(p_Tin IN VARCHAR2
    					    , p_Tmt_Ma_Muc IN VARCHAR2
    						, p_Tmt_Ma_TMuc IN VARCHAR2
    						, p_Tmt_Ma_Thue IN VARCHAR2
    						, p_KyKK IN DATE ) RETURN qlt_xltk_gdich.so_tien%TYPE IS
	CURSOR c_So_Tien IS
		SELECT SUM(So_Tien)
		FROM(SELECT SUM(NVL(so_tien,0))so_tien
			 FROM qlt_xltk_gdich
     		 WHERE tin = p_Tin
   		     AND tmt_ma_muc = p_Tmt_Ma_Muc
    		 AND tmt_ma_tmuc = p_Tmt_Ma_TMuc
			 AND tmt_ma_thue = p_Tmt_Ma_Thue
			 AND kykk_tu_ngay >= Trunc(p_KyKK, 'RRRR')
			 AND kykk_den_ngay <= Last_Day (Add_Months(p_KyKK,11))
			 AND dgd_ma_gdich = '94'
			 AND dgd_kieu_gdich ='25'
		UNION ALL
			 SELECT SUM(decode(dgd_ma_gdich,58,(-1)*NVL(so_tien,0),NVL(so_tien,0))) so_tien
			 FROM qlt_xltk_gdich gdich
                , qlt_tkhai_hdr hdr
    		 WHERE gdich.tin = p_Tin
 			 AND gdich.hdr_Id = hdr.Id
 			 AND hdr.ltd = 0
			 AND hdr.dtk_ma_loai_tkhai IN ('04','24')
			 AND gdich.tmt_ma_muc = p_Tmt_Ma_Muc
			 AND gdich.tmt_ma_tmuc = p_Tmt_Ma_TMuc
			 AND gdich.tmt_ma_thue = p_Tmt_Ma_Thue
			 AND gdich.kykk_tu_ngay >= Trunc(p_KyKK, 'RRRR')
			 AND gdich.kykk_den_ngay <= Last_Day (Add_Months(p_KyKK,11))
			 AND ((gdich.dgd_ma_gdich IN ('06','12','18')
			     AND gdich.dgd_kieu_gdich = '21')
				 OR (gdich.dgd_ma_gdich = '58' AND gdich.dgd_kieu_gdich = '25'))
		UNION ALL
			SELECT SUM(NVL(so_tien,0)) so_tien
			FROM qlt_xltk_gdich gdich
               , qlt_ds_an_dinh_hdr hdr
			WHERE gdich.tin = p_Tin
			AND gdich.hdr_Id = hdr.Id
			AND hdr.dtk_ma IN ('04','24')
			AND	gdich.tmt_ma_muc = p_Tmt_Ma_Muc
		    AND gdich.tmt_ma_tmuc = p_Tmt_Ma_TMuc
		    AND gdich.tmt_ma_thue = p_Tmt_Ma_Thue
 		    AND gdich.kykk_tu_ngay >= Trunc(p_KyKK, 'RRRR')
			AND gdich.kykk_den_ngay <= Last_Day (Add_Months(p_KyKK,11))
		    AND gdich.dgd_ma_gdich IN ('22','24')
		    AND gdich.dgd_kieu_gdich = '21');
	v_So_Tien	qlt_xltk_gdich.so_tien%TYPE;
BEGIN
    v_So_Tien := 0;
    OPEN c_So_Tien;
    FETCH c_So_Tien INTO v_So_Tien;
    CLOSE c_So_Tien;
    RETURN NVL(v_So_Tien,0);
END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 13/04/2006
Muc dich: Thuc hien lay so ke khai thue GTGT truc tiep cua DTNT trong 12 thang
Tham so:
        - p_Tin: Ma so thue cua DTNT
        - p_Tmt_Ma_Muc: Ma muc
        - p_Tmt_Ma_TMuc: Ma tieu muc
        - p_Tmt_Ma_Thue: Ma thue
        - p_KyKK: Ky ke khai
*******************************************************************************/
    FUNCTION Fnc_So_KKhai_GTGT_TT(p_Tin VARCHAR2,
							      p_Tmt_Ma_Muc VARCHAR2,
						          p_Tmt_Ma_TMuc VARCHAR2,
					  		      p_Tmt_Ma_Thue VARCHAR2,
							      p_KyKK DATE ) RETURN NUMBER IS
    CURSOR c_So_Tien IS
        SELECT SUM(so_tien)
		FROM(
				SELECT SUM(NVL(so_tien,0)) so_tien
				FROM qlt_xltk_gdich
				WHERE (tin = p_Tin)
				  AND (tmt_ma_muc = p_Tmt_Ma_Muc)
				  AND (tmt_ma_tmuc = p_Tmt_Ma_TMuc)
				  AND (tmt_ma_thue = p_Tmt_Ma_Thue)
				  AND (kykk_tu_ngay >= TRUNC(p_KyKK, 'RRRR'))
				  AND (kykk_den_ngay <= LAST_DAY(ADD_MONTHS(p_KyKK,11)))
				  AND (dgd_ma_gdich IN ('94'))
				  AND (dgd_kieu_gdich IN ('25'))
  	            UNION ALL
				SELECT SUM(DECODE(dgd_ma_gdich,58,(-1)*NVL(so_tien,0),NVL(so_tien,0))) so_tien
				FROM qlt_xltk_gdich gdich, qlt_tkhai_hdr hdr
				WHERE (gdich.tin = p_Tin)
				  AND (gdich.hdr_id = hdr.Id)
				  AND (hdr.ltd = 0)
				  AND (hdr.dtk_ma_loai_tkhai = '02')
				  AND (gdich.tmt_ma_muc = p_Tmt_Ma_Muc)
				  AND (gdich.tmt_ma_tmuc = p_Tmt_Ma_TMuc)
				  AND (gdich.tmt_ma_thue = p_Tmt_Ma_Thue)
				  AND (gdich.kykk_tu_ngay >= TRUNC(p_KyKK, 'RRRR'))
				  AND (gdich.kykk_den_ngay <= LAST_DAY (ADD_MONTHS(p_KyKK,11)))
				  AND ((gdich.dgd_ma_gdich IN ('06','12','18')
				  AND gdich.dgd_kieu_gdich = '21')
				  			OR (gdich.dgd_ma_gdich = '58' AND gdich.dgd_kieu_gdich = '25'))
		        UNION ALL
				--SonLH: Chi lay so an dinh tren QToan thue GTGT truc tiep
				SELECT SUM(NVL(so_tien,0)) so_tien
				FROM qlt_xltk_gdich gdich, qlt_ds_an_dinh_hdr hdr
				WHERE (gdich.TIN = p_Tin)
				  And (gdich.Hdr_Id = Hdr.Id)
				  And (hdr.Dtk_Ma = '02')
				  And (gdich.tmt_ma_muc = p_Tmt_Ma_Muc)
				  And (gdich.tmt_ma_tmuc = p_Tmt_Ma_TMuc)
				  And (gdich.tmt_ma_thue = p_Tmt_Ma_Thue)
				  And (gdich.kykk_tu_ngay >= TRUNC(p_KyKK, 'RRRR'))
				  And (gdich.kykk_den_ngay <= LAST_DAY (ADD_MONTHS(p_KyKK,11)))
				  And (gdich.dgd_ma_gdich IN ('22','24'))
				  And (gdich.dgd_kieu_gdich = '21'));

    v_So_Tien NUMBER(20,2) := 0;

    BEGIN
        OPEN c_So_Tien;
        FETCH c_So_Tien INTO v_So_Tien;
        CLOSE c_So_Tien;
        RETURN NVL(v_So_Tien,0);
    END;
/*******************************************************************************
Nguoi lap: Nguyen Hoang Gia Khai
Ngay lap: 25/05/2006
Muc dich: Thuc hien lay thong tin cua DTNT
Tham so:
        - p_Date: Luu tru ngay hien tai
        - p_Ma: Ma ho so
*******************************************************************************/
    FUNCTION Fnc_Get_So_Hoso(p_Date IN DATE,
                             p_Ma IN VARCHAR2) RETURN VARCHAR2 IS

    	v_Tmp	    VARCHAR2(20);
    	v_So_Hoso	VARCHAR2(30);
    	v_Id 		NUMBER := 0;
    	--Lay so ho so lon nhat(phan sau /)
    	CURSOR c_Count_Hs(v_so VARCHAR2) IS
    		   SELECT MAX(TO_NUMBER(SUBSTR(so_hoso,INSTR(so_hoso,'/',-1,1)+1)))
               FROM qhs_so_hoso
               WHERE SUBSTR(so_hoso,1,2) = v_so;
    	v_Sohs VARCHAR2(10);
    	v_Num NUMBER(20) := 0;
    BEGIN
    	v_Tmp := TO_CHAR(p_Date,'RR')
     	      		   ||TO_CHAR(p_Date,'MM')
     				   ||TO_CHAR(p_Date,'DD')
    				   ||'/'||p_Ma||'/';

    	v_sohs := TO_CHAR(p_Date,'RR');

    	OPEN c_Count_Hs(v_Sohs);
    	FETCH c_Count_Hs INTO v_Num;
    	CLOSE c_Count_Hs;
    	IF(v_Num IS NULL) THEN
    		v_Num := 0;
    	END IF;
    	v_Num := v_Num + 1;
    	v_So_Hoso := v_Tmp || TO_CHAR(v_num);
    	RETURN v_So_Hoso;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 12/01/2006
Muc dich: Thuc hien lay thong tin cua DTNT
Tham so:
        - p_Tin: Ma so thue
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
*******************************************************************************/
    PROCEDURE Prc_Thong_Tin_Dtnt(p_Tin VARCHAR2,
                                 p_Record_Dtnt OUT Record_Dtnt) IS
        CURSOR c_Dtnt IS
            SELECT *
            FROM rcv_v_dtnt dtnt
            WHERE (dtnt.tin = p_Tin);
        vc_Dtnt c_Dtnt%ROWTYPE;
    BEGIN
        OPEN c_Dtnt;
        FETCH c_Dtnt INTO vc_Dtnt;
        IF (c_Dtnt%FOUND) THEN
            p_Record_Dtnt.tin := vc_Dtnt.tin;
            p_Record_Dtnt.ten_dtnt := vc_Dtnt.ten_dtnt;
            p_Record_Dtnt.dia_chi := vc_Dtnt.dia_chi;
            p_Record_Dtnt.dien_thoai := vc_Dtnt.dien_thoai;
            p_Record_Dtnt.fax := vc_Dtnt.fax;
            p_Record_Dtnt.email := vc_Dtnt.email;
            p_Record_Dtnt.ma_cqt := vc_Dtnt.ma_cqt;
            p_Record_Dtnt.ma_tinh := vc_Dtnt.ma_tinh;
            p_Record_Dtnt.ma_huyen := vc_Dtnt.ma_huyen;
            p_Record_Dtnt.ma_phong := vc_Dtnt.ma_phong;
            p_Record_Dtnt.ma_canbo := vc_Dtnt.ma_canbo;
            p_Record_Dtnt.ma_cap := vc_Dtnt.ma_cap;
            p_Record_Dtnt.ma_chuong := vc_Dtnt.ma_chuong;
            p_Record_Dtnt.ma_loai := vc_Dtnt.ma_loai;
            p_Record_Dtnt.ma_khoan := vc_Dtnt.ma_khoan;
            p_Record_Dtnt.ngay_tchinh := vc_Dtnt.ngay_tchinh;
        END IF;
        CLOSE c_Dtnt;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 09/02/2006
Muc dich: Thuc hien lay han nop cua mot to khai hoac quyet toan
Tham so:
        - p_Ma_TKhai: Ma to khai
        - p_Loai: Loai to khai hay quyet toan
        - p_Cuoi_Ky: Ngay cuoi cua ky ke khai ung voi
                     tung loai to khai hay quyet toan
        - p_Nam_TChinh: Nam tai chinh cua DTNT
*******************************************************************************/
    FUNCTION Fnc_Han_Nop(p_Ma_TKhai VARCHAR2,
                         p_Loai VARCHAR2,
                         p_Cuoi_Ky DATE,
                         p_Nam_TChinh DATE DEFAULT NULL) RETURN DATE IS
        --Kiem tra co theo nam tai chinh hay khong
        CURSOR c_Nam_TChinh IS
            SELECT gia_tri
            FROM rcv_thamso
            WHERE (ten = 'THEO_NAM_TAICHINH');
        v_Nam_TChinh NUMBER;

        --Lay cac loai to khai tinh toan theo nam tai chinh
        CURSOR c_TKhai_TChinh IS
            SELECT gia_tri
            FROM rcv_thamso
            WHERE (ten = 'LOAI_TK_TAICHINH');
        v_TKhai_TChinh VARCHAR2(1000);

        --Kiem tra to khai dang xu ly co tinh toan theo nam tai chinh khong
        CURSOR c_TKhai_TChinh_Qlt IS
            SELECT ma_tkhai_qlt
            FROM rcv_map_tkhai
            WHERE INSTR(v_TKhai_TChinh, ma_tkhai) > 0
              AND (ma_tkhai_qlt = p_Ma_TKhai);

        --Tinh han nop to khai khong theo nam tai chinh
        CURSOR c_HanNop_TKhai IS
            SELECT ADD_MONTHS(p_Cuoi_Ky,1) han_nop_tkhai
            FROM qlt_dm_tkhai
            WHERE (ma = p_Ma_TKhai);

        --Tinh han nop quyet toan khong theo nam tai chinh
        CURSOR c_HanNop_QToan IS
            SELECT ADD_MONTHS(p_Cuoi_Ky,TO_NUMBER(thang)) han_nop_qtoan
            FROM qlt_dm_qtoan
            WHERE (ma = p_Ma_TKhai);

        --Tinh han nop to khai theo nam tai chinh
        CURSOR c_HanNop_TKhai_TChinh IS
            SELECT DECODE(kieu_ky,'M',ADD_MONTHS(p_Cuoi_Ky,1)
                                 ,'Q',ADD_MONTHS(TRUNC(LAST_DAY(TO_DATE(TO_CHAR(p_Nam_TChinh,'DD/MM')||'/'||TO_CHAR(p_Cuoi_Ky,'RRRR'),'DD/MM/RRRR'))),TO_NUMBER(TO_CHAR(p_Cuoi_Ky,'Q')) * 3)
                                 ,'Y',ADD_MONTHS(TRUNC(LAST_DAY(TO_DATE(TO_CHAR(p_Nam_TChinh,'DD/MM')||'/'||TO_CHAR(p_Cuoi_Ky,'RRRR'),'DD/MM/RRRR'))),12)) han_nop_tkhai
            FROM qlt_dm_tkhai
            WHERE (ma = p_Ma_TKhai);

        --Tinh han nop quyet toan theo nam tai chinh
        CURSOR c_HanNop_QToan_TChinh IS
            SELECT ADD_MONTHS(LAST_DAY(TO_DATE(TO_CHAR(p_Nam_TChinh,'DD/MM')||'/'||TO_CHAR(p_Cuoi_Ky,'RRRR'),'DD/MM/RRRR')),TO_NUMBER(thang) + 11) han_nop_qtoan
            FROM qlt_dm_qtoan
            WHERE (ma = p_Ma_TKhai);

        v_HanNop DATE;
        v_TKhai_Khong_Nam_TChinh VARCHAR2(2);

    BEGIN
        OPEN c_Nam_TChinh;
        FETCH c_Nam_TChinh INTO v_Nam_TChinh;
        CLOSE c_Nam_TChinh;

        IF (v_Nam_TChinh = 1) THEN --Theo nam tai chinh
            IF (p_Loai = 'TK') THEN
                OPEN c_HanNop_TKhai;
                FETCH c_HanNop_TKhai INTO v_HanNop;
                CLOSE c_HanNop_TKhai;
            ELSIF (p_Loai = 'QT') THEN
                OPEN c_HanNop_QToan;
                FETCH c_HanNop_QToan INTO v_HanNop;
                CLOSE c_HanNop_QToan;
            END IF;
        ELSIF (v_Nam_TChinh = 0) THEN --Khong theo nam tai chinh
            OPEN c_TKhai_TChinh;
            FETCH c_TKhai_TChinh INTO v_TKhai_TChinh;
            CLOSE c_TKhai_TChinh;

            FOR vc_TKhai_TChinh_Qlt IN c_TKhai_TChinh_Qlt LOOP
                IF (p_Loai = 'TK') THEN
                    OPEN c_HanNop_TKhai_TChinh;
                    FETCH c_HanNop_TKhai_TChinh INTO v_HanNop;
                    CLOSE c_HanNop_TKhai_TChinh;
                ELSIF (p_Loai = 'QT') THEN
                    OPEN c_HanNop_QToan_TChinh;
                    FETCH c_HanNop_QToan_TChinh INTO v_HanNop;
                    CLOSE c_HanNop_QToan_TChinh;
                END IF;
            END LOOP;

            SELECT ma_tkhai INTO v_TKhai_Khong_Nam_TChinh
            FROM rcv_map_tkhai
            WHERE (ma_tkhai_qlt = p_Ma_TKhai)
              AND (loai = p_Loai);

            IF (INSTR(v_TKhai_TChinh,v_TKhai_Khong_Nam_TChinh) = 0) THEN
                IF (p_Loai = 'TK') THEN
                    OPEN c_HanNop_TKhai;
                    FETCH c_HanNop_TKhai INTO v_HanNop;
                    CLOSE c_HanNop_TKhai;
                ELSIF (p_Loai = 'QT') THEN
                    OPEN c_HanNop_QToan;
                    FETCH c_HanNop_QToan INTO v_HanNop;
                    CLOSE c_HanNop_QToan;
                END IF;
            END IF;
        END IF;

        RETURN v_HanNop;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 09/04/2006
Muc dich: Kiem tra sai so cua cac chi tieu tren to khai GTGT truc tiep
Tham so:
        - p_So_DTNT: So cua DTNT
        - p_So_CQT: so cua CQT tinh
        - p_KyLB: Ky lap bo cua to khai
*******************************************************************************/
    FUNCTION Fnc_Ktra_Ctieu(p_So_DTNT NUMBER,
                            p_So_CQT  NUMBER,
                            p_KyLB DATE) RETURN BOOLEAN IS
        CURSOR c_Gia_tri(p_Loai VARCHAR2) IS
            SELECT NVL(gia_tri,0) gia_tri
            FROM qlt_dkien_ktra_tkhai
            WHERE (lth_ma_thue = '01')
              AND (loai_dkien = p_Loai)
              AND (tu_ky <= p_KyLB)
              AND (NVL(den_ky, TO_DATE('16/11/2050','DD/MM/YYYY')) >= p_KyLB);

        v_Gia_tri c_Gia_tri%ROWTYPE;
        v_So_DTNT Number;
        v_So_Cqt NUMBER;
        v_Return BOOLEAN;
        v_Return1 BOOLEAN;
        v_Tyle BOOLEAN;
        v_GTTD BOOLEAN;
        v_Gt_Tyle NUMBER;
        v_Gt_GTTD NUMBER;
    BEGIN
    	v_So_DTNT := p_So_DTNT;
        v_So_CQT := p_So_CQT;
        IF (v_So_DTNT = v_So_CQT) THEN
    		RETURN TRUE;
    	END IF;
        --khainhg kiem tra so DTNT, CQT
        IF(v_So_DTNT >0) AND (v_So_CQT <0) THEN
            RETURN FALSE;
        END IF;
        IF(v_So_DTNT <0) AND (v_So_CQT >0) THEN
    		RETURN FALSE;
       	END IF;
        IF(v_So_DTNT <0) AND (v_So_CQT <0) THEN
       		v_So_DTNT := Abs(v_So_DTNT);
       		v_So_CQT := Abs(v_So_CQT);
       	END IF;
        IF NVL(v_So_DTNT,0) =0 AND NVL(v_So_CQT,0) <>0 THEN
       		RETURN FALSE;
       	END IF;
        IF NVL(v_So_DTNT,0) <>0 AND NVL(v_So_CQT,0) =0 THEN
       		RETURN FALSE;
       	END IF;
        --End Khainhg
    	OPEN c_Gia_tri('0');
    	FETCH c_Gia_tri INTO v_Gia_tri;
    	IF (c_Gia_tri%NOTFOUND) THEN
    	    CLOSE c_Gia_tri;
    	    v_Tyle := FALSE;
    	ELSE
    	    v_Gt_Tyle := v_Gia_tri.Gia_tri;
    	    v_Tyle := TRUE;
    	    CLOSE c_Gia_tri;
    	END IF;

    	OPEN c_Gia_tri('1');
    	FETCH c_Gia_tri INTO v_Gia_tri;
        IF (c_Gia_tri%NOTFOUND) THEN
          CLOSE c_Gia_tri;
          v_GTTD := FALSE;
        ELSE
          v_Gt_GTTD := v_Gia_tri.Gia_tri;
          CLOSE c_Gia_tri;
          v_GTTD := TRUE;
        END IF;
        IF (v_Tyle AND v_GTTD) THEN --Co ca 2 tham so
            IF (v_So_CQT <> 0) THEN
        		IF (ABS(v_So_DTNT - v_So_CQT)*100/v_So_CQT > v_Gt_Tyle) THEN
        			RETURN FALSE;
        		END IF;
            END IF;
        	IF (ABS(v_So_DTNT - v_So_CQT) > v_Gt_GTTD) THEN
        		RETURN FALSE;
        	END IF;
        	RETURN TRUE;
        END IF;

        IF (v_Tyle AND NOT(v_GTTD)) THEN -- Chi co tham so ty le
            IF (v_So_CQT <> 0) THEN
            		IF (ABS(v_So_DTNT - v_So_CQT)*100/v_So_CQT > v_Gt_Tyle) THEN
            			RETURN FALSE;
            		ELSE
            			RETURN TRUE;
            		END IF;
            END IF;
        END IF;
        IF (v_Tyle = FALSE AND v_GTTD = TRUE) THEN -- Chi co tham so GTTD
        	IF (ABS(v_So_DTNT - v_So_CQT) > v_Gt_GTTD) THEN
        		RETURN FALSE;
        	ELSE
        		RETURN TRUE;
        	END IF;
        END IF;

        IF (v_Tyle = FALSE AND v_GTTD = FALSE) THEN -- Khong co tham so nao
            IF (v_So_DTNT <> v_So_CQT) THEN
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
Muc dich: Thuc hien do du lieu to khai GTGT khau tru vao CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai
*******************************************************************************/
    PROCEDURE Prc_TKhai_GTGT(p_Record_Of_Header Record_Hdr,
                             p_Record_Dtnt Record_Dtnt,
                             p_TKhai_Exits_Id NUMBER,
                             p_Tthai_Tkhai VARCHAR2) IS
	    TYPE Record_Dtl_GTGT IS RECORD(id NUMBER(10,0),
                                  ctk_id NUMBER(10,0),
                                  so_tt NUMBER(3,0),
                                  doanhso_dtnt NUMBER(20,2),
                                  sothue_dtnt NUMBER(20,2),
                                  doanhso_cqt NUMBER(20,2),
                                  sothue_cqt NUMBER(20,2),
                                  ke_khai_sai VARCHAR2(1),
                                  ma_so_ct_thue VARCHAR2(5),
                                  ma_so_ct_doanhso VARCHAR2(5));
        TYPE Array_Of_Record_Dtl IS TABLE OF Record_Dtl_GTGT INDEX BY BINARY_INTEGER;
        v_Array_Of_Record_Dtl Array_Of_Record_Dtl;

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

        v_Thue_Dau_Ra NUMBER(20,2);
        v_Thue_Dau_Vao NUMBER(20,2);
        v_Thue_KTru_KySau NUMBER(20,2);
        v_Thue_PSinh_KyNay NUMBER(20,2);
        v_Thue_KTru_KyNay NUMBER(20,2);
        v_Thue_PNop_KyNay NUMBER(20,2);
        v_Hdr_Id NUMBER(10,0);
        v_So_KTru NUMBER(20,2);
        v_PSinh_TKy NUMBER(20,2);
        v_Tthai_Tkhai VARCHAR2(1);
        v_Count NUMBER(10) := 0;
        v_Index NUMBER(10) := 0;
        v_Temp NUMBER(20,2) := 0;
        v_Nguong_Min NUMBER;
        v_Sai_SoHoc VARCHAR2(1) := NULL;
        v_HanNop DATE;
        --Bien luu ds phu luc dung cho ghi so nhan
        v_Ds_Pluc VARCHAR2(10) := NULL;

    BEGIN
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

       	--Chi Tieu [30]*5% - Nguong <= |[31]| <= [30]*5% + Nguong
        v_Temp := ROUND((v_Array_Of_Record_Dtl(18).doanhso_cqt*5)/100);
        v_Nguong_Min := v_Temp/1000; --Sai so 0.1%

        IF (v_Nguong_Min > 1000000) THEN
            v_Nguong_Min := 1000000;
        END IF;

        IF ((ABS(v_Array_Of_Record_Dtl(18).sothue_cqt) < v_Temp - v_Nguong_Min) OR
            (ABS(v_Array_Of_Record_Dtl(18).sothue_cqt) > v_Temp + v_Nguong_Min)) THEN
            v_Array_Of_Record_Dtl(18).ke_khai_sai := 'Y';
            v_Array_Of_Record_Dtl(18).sothue_cqt := v_Temp;
        END IF;

       	--Chi Tieu [32]*10% - Nguong <= |[33]| <= [32]*10% + Nguong
        v_Temp := ROUND((v_Array_Of_Record_Dtl(19).doanhso_cqt*10)/100);
        v_Nguong_Min := v_Temp/1000; --Sai so 0.1%
        IF (v_Nguong_Min > 1000000) THEN
            v_Nguong_Min := 1000000;
        END IF;

        IF ((ABS(v_Array_Of_Record_Dtl(19).sothue_cqt) < v_Temp - v_Nguong_Min) OR
            (ABS(v_Array_Of_Record_Dtl(19).sothue_cqt) > v_Temp + v_Nguong_Min)) THEN
            v_Array_Of_Record_Dtl(19).ke_khai_sai := 'Y';
            v_Array_Of_Record_Dtl(19).sothue_cqt := v_Temp;
        END IF;

        /* Xu ly chi tieu 28 = 31 + 33 */
        v_Array_Of_Record_Dtl(16).sothue_cqt := NVL(v_Array_Of_Record_Dtl(18).sothue_cqt,0) +
                                                NVL(v_Array_Of_Record_Dtl(19).sothue_cqt,0);

        /* Xu ly chi tieu 25 = 28*/
        v_Array_Of_Record_Dtl(14).sothue_cqt := NVL(v_Array_Of_Record_Dtl(16).sothue_cqt,0);

        /* Xu ly chi tieu 39 = 25 + 35 - 37 */
        v_Array_Of_Record_Dtl(23).sothue_cqt := NVL(v_Array_Of_Record_Dtl(14).sothue_cqt,0) +
                                                 NVL(v_Array_Of_Record_Dtl(21).sothue_cqt,0) -
                                                 NVL(v_Array_Of_Record_Dtl(22).sothue_cqt,0);

        /* 39 - 23 - 11 */
        v_Temp := NVL(v_Array_Of_Record_Dtl(23).sothue_cqt,0) -
                  NVL(v_Array_Of_Record_Dtl(12).sothue_dtnt,0) -
                  NVL(v_Array_Of_Record_Dtl(2).sothue_cqt,0);

        IF (v_Temp > 0) THEN
            --Neu > 0 ghi vao chi tieu [40]
            v_Array_Of_Record_Dtl(25).sothue_cqt := v_Temp;
        ELSIF (v_Temp < 0) THEN
            --Neu < 0 ghi vao chi tieu [41]
            v_Array_Of_Record_Dtl(26).sothue_cqt := ABS(v_Temp);
        ELSE
            v_Array_Of_Record_Dtl(25).sothue_cqt := 0;
            v_Array_Of_Record_Dtl(26).sothue_cqt := 0;
        END IF;

        /*Xu ly chi tieu [43] = [41]-[42], khi khong co chi tieu [40]*/
        IF (v_Array_Of_Record_Dtl(25).sothue_cqt = 0) THEN
            v_Array_Of_Record_Dtl(28).sothue_cqt := NVL(v_Array_Of_Record_Dtl(26).sothue_cqt,0) -
                                                    NVL(v_Array_Of_Record_Dtl(27).sothue_cqt,0);
        END IF;

        FOR i IN 1..v_Count LOOP
            IF (i NOT IN (18,19)) THEN
                IF (v_Array_Of_Record_Dtl(i).doanhso_dtnt <>
                    v_Array_Of_Record_Dtl(i).doanhso_cqt
                    OR
                    v_Array_Of_Record_Dtl(i).sothue_dtnt <>
                    v_Array_Of_Record_Dtl(i).sothue_cqt)
                THEN
                    v_Array_Of_Record_Dtl(i).ke_khai_sai := 'Y';
                    v_Sai_SoHoc := 'Y';
                END IF;
            END IF;
        END LOOP;

        /*Tinh toan so thue phat sinh*/
        v_Thue_Dau_Ra := NVL(v_Array_Of_Record_Dtl(23).sothue_dtnt,0); --Chi tieu 39
        v_Thue_Dau_Vao := NVL(v_Array_Of_Record_Dtl(12).sothue_dtnt,0); --Chi tieu 23
        v_Thue_KTru_KySau := NVL(v_Array_Of_Record_Dtl(28).sothue_dtnt,0); --Chi tieu 43;
        v_Thue_PSinh_KyNay := NVL(v_Array_Of_Record_Dtl(23).sothue_dtnt,0) -
                              NVL(v_Array_Of_Record_Dtl(12).sothue_dtnt,0); -- = 39-23
        v_Thue_KTru_KyNay := NVL(v_Array_Of_Record_Dtl(2).sothue_dtnt,0) +
                             NVL(v_Array_Of_Record_Dtl(12).sothue_dtnt,0) -
                             NVL(v_Array_Of_Record_Dtl(27).sothue_dtnt,0) -
                             NVL(v_Array_Of_Record_Dtl(28).sothue_dtnt,0); -- = 11+23-42-43
        v_Thue_PNop_KyNay := NVL(v_Array_Of_Record_Dtl(25).sothue_dtnt,0); --Chi tieu 40;
        /*Ket thuc tinh toan so thue phat sinh*/

        /********************************************************/
        /*Ket thuc kiem tra ke khai sai voi cac chi tieu to khai*/

        /*Xu ly to khai*/
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            /*Neu da ton tai to khai trong ky ke khai*/
            Qlt_Pck_Gdich.Prc_Lay_Thamso(p_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Sinh giao dich*/
            Qlt_Pck_Gdich.Prc_Set_GTGT_2004;

            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR','14',p_TKhai_Exits_Id);

            /*Thong tin Header*/
            UPDATE qlt_tkhai_hdr
               SET co_loi = v_Sai_SoHoc,
                   co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   co_gtrinh_02a = p_Record_Of_Header.co_gtrinh_02a,
                   co_gtrinh_02b = p_Record_Of_Header.co_gtrinh_02b,
                   co_gtrinh_02c = p_Record_Of_Header.co_gtrinh_02c,
                   --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = '4', --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat
            WHERE (id = p_TKhai_Exits_Id)
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
                WHERE (tkh_id = p_TKhai_Exits_Id)
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
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

            /*Thong tin phu luc 2A*/
            /*Xoa ban ghi cu neu co*/
            DELETE FROM qlt_gtrinh_gtgt_kt_02a
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

            IF (p_Record_Of_Header.co_gtrinh_02a = 'Y') THEN
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
                       p_TKhai_Exits_Id,
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
            END IF;

            /*Xoa ban ghi cu neu co cua phu luc 2B, 2C neu co*/
            DELETE FROM qlt_gtrinh_gtgt_kt_02bc
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

            /*Insert ban ghi phu luc 2B moi*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02bc(id,
                                                tkh_id,
                                                tkh_ltd,
                                                ctg_id,
                                                gia_tri_kkhai)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   p_TKhai_Exits_Id,
                   0,
                   pluc2b.ctg_id,
                   pluc2b.gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2b pluc2b
            WHERE (pluc2b.hdr_id = p_Record_Of_Header.id);

            /*Insert ban ghi phu luc 2C moi*/
            INSERT INTO qlt_gtrinh_gtgt_kt_02bc(id,
                                                tkh_id,
                                                tkh_ltd,
                                                ctg_id,
                                                gia_tri_kkhai)
            SELECT qlt_dm_ctieu_tkhai_seq.NEXTVAL,
                   p_TKhai_Exits_Id,
                   0,
                   pluc2c.ctg_id,
                   pluc2c.gia_tri_ctieu
            FROM rcv_v_tkhai_gtgt_kt_pluc2c pluc2c
            WHERE (pluc2c.hdr_id = p_Record_Of_Header.id);
            /*Ket thuc cap nhat thong tin cho to khai thay the*/
        ELSE
            /*Neu chua ton tai to khai trong ky ke khai*/
            IF (Fnc_AnDinh_Exits(p_Record_Of_Header.tin,
                                 p_Record_Of_Header.loai_tkhai,
                                 p_Record_Of_Header.kykk_tu_ngay,
                                 p_Record_Of_Header.kykk_den_ngay)) THEN /*Neu co an dinh*/
                v_Tthai_Tkhai := '3'; /*To khai nop cham*/
            ELSE /*Neu chua co an dinh*/
                v_HanNop := Fnc_Han_Nop(p_Record_Of_Header.loai_tkhai,
                                        'TK',
                                        p_Record_Of_Header.kykk_den_ngay);
                IF (p_Record_Of_Header.ngay_nop <= v_HanNop) THEN
                    --To khai dung han
                    v_Tthai_Tkhai := '1'; /*To khai chinh thuc*/
                ELSE
                    --To khai khong dung han
                    v_Tthai_Tkhai := '3'; /*To khai nop cham*/
                END IF;
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
                                      co_loi,
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
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   v_Sai_SoHoc,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   p_Record_Of_Header.co_gtrinh_02a,
                   p_Record_Of_Header.co_gtrinh_02b,
                   p_Record_Of_Header.co_gtrinh_02c,
                   NULL,  --p_Record_Of_Header.so_hieu_tep,
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
                    p_Record_Dtnt.ma_cap,
                    p_Record_Dtnt.ma_chuong,
                    p_Record_Dtnt.ma_loai,
                    p_Record_Dtnt.ma_khoan,
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
        --Xoa mang trung gian
        v_Array_Of_Record_Dtl.DELETE;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);

        /*Ghi so ho so nhan*/
        IF p_Record_Of_Header.co_gtrinh_02a IS NOT NULL THEN
            v_Ds_Pluc := v_Ds_Pluc||'77';
        END IF;
        IF p_Record_Of_Header.co_gtrinh_02b IS NOT NULL THEN
            v_Ds_Pluc := v_Ds_Pluc||'78';
        END IF;
        IF p_Record_Of_Header.co_gtrinh_02c IS NOT NULL THEN
            v_Ds_Pluc := v_Ds_Pluc||'79';
        END IF;

        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '14',--To khai GTGT KT
                         '02',--To khai nop thue
                         v_Ds_Pluc);
    EXCEPTION
    WHEN OTHERS THEN
        ROLLBACK;
        QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_GTGT'
                                    , FALSE
                                    , NULL);
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 12/01/2006
Muc dich: Thuc hien do du lieu to khai TNDN quy vao cac CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai
*******************************************************************************/
    PROCEDURE Prc_TKhai_TNDN_Quy(p_Record_Of_Header Record_Hdr,
                                 p_Record_Dtnt Record_Dtnt,
                                 p_TKhai_Exits_Id NUMBER,
                                 p_Tthai_Tkhai VARCHAR2) IS
        CURSOR c_TKhai_TNDN_Quy IS
            SELECT *
            FROM rcv_v_tkhai_tndn_quy tkhai_quy
            WHERE (tkhai_quy.hdr_id = p_Record_Of_Header.id)
            ORDER BY tkhai_quy.so_tt;

        CURSOR c_Ma_Thue IS
            SELECT dm.lte_ma_thue
            FROM qlt_dm_tkhai dm
            WHERE (dm.ma = p_Record_Of_Header.loai_tkhai);

        v_Hdr_Id NUMBER(10);
        v_Ma_Thue VARCHAR2(2);
        v_Thue_PSinh NUMBER(20,2);
        v_Tthai_Tkhai VARCHAR2(1);
        v_HanNop DATE;

    BEGIN
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
        --Neu da ton tai to khai trong ky ke khai
            --Gan tham so sinh giao dich
            Qlt_Pck_Gdich.Prc_Lay_Thamso(p_Tthai_Tkhai, p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            --Thuc hien backup to khai
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);

                   --Update thong tin Header
            UPDATE qlt_tkhai_hdr
               SET co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = '4', --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat
            WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);

            --Update du lieu to khai TNDN quy
            FOR vc_TKhai_TNDN_Quy IN c_TKhai_TNDN_Quy LOOP
                UPDATE qlt_tkhai_tndn_quy
                SET so_dtnt = NVL(vc_TKhai_TNDN_Quy.so_dtnt,0),
                    so_cqt = NVL(vc_TKhai_TNDN_Quy.so_dtnt,0),
                    ke_khai_sai = NULL
                WHERE (tkh_id = p_TKhai_Exits_Id)
                  AND (tkh_ltd = 0)
                  AND (so_tt = vc_TKhai_TNDN_Quy.so_tt);

                --Update so thue phat sinh trong bang phat sinh
                IF (vc_TKhai_TNDN_Quy.so_tt = 9) THEN
                    UPDATE qlt_psinh_tkhai
                    SET thue_psinh = NVL(vc_TKhai_TNDN_Quy.so_dtnt,0)
                    WHERE (tkh_id = p_TKhai_Exits_Id)
                      AND (tkh_ltd = 0);
                END IF;
            END LOOP;
        ELSE
            --Neu chua ton tai to khai trong ky ke khai
            IF (Fnc_AnDinh_Exits(p_Record_Of_Header.tin,
                                 p_Record_Of_Header.loai_tkhai,
                                 p_Record_Of_Header.kykk_tu_ngay,
                                 p_Record_Of_Header.kykk_den_ngay)) THEN /*Neu co an dinh*/
                v_Tthai_Tkhai := '3'; /*To khai nop cham*/
            ELSE /*Neu chua co an dinh*/
                v_HanNop := Fnc_Han_Nop(p_Record_Of_Header.loai_tkhai,
                                        'TK',
                                        p_Record_Of_Header.kykk_den_ngay,
                                        p_Record_Dtnt.ngay_tchinh);
                IF (p_Record_Of_Header.ngay_nop <= v_HanNop) THEN
                    --To khai dung han
                    v_Tthai_Tkhai := '1'; /*To khai chinh thuc*/
                ELSE
                    --To khai khong dung han
                    v_Tthai_Tkhai := '3'; /*To khai nop cham*/
                END IF;
            END IF;

            --Gan tham so sinh giao dich
            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai, p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);

            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Xu ly du lieu to khai Header
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
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   NULL, --p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat);

            --Xu ly du lieu to khai Detail
            FOR vc_TKhai_TNDN_Quy IN c_TKhai_TNDN_Quy LOOP
                INSERT INTO qlt_tkhai_tndn_quy (id,
                                                tkh_id,
                                                tkh_ltd,
                                                ctk_id,
                                                so_tt,
                                                so_dtnt,
                                                so_cqt,
                                                ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        vc_TKhai_TNDN_Quy.ctk_id,
                        NVL(vc_TKhai_TNDN_Quy.so_tt,0),
                        NVL(vc_TKhai_TNDN_Quy.so_dtnt,0),
                        NVL(vc_TKhai_TNDN_Quy.so_dtnt,0),
                        null);
                --Luu so thue phat sinh
                IF (vc_TKhai_TNDN_Quy.so_tt = 9) THEN
                    v_Thue_PSinh := NVL(vc_TKhai_TNDN_Quy.so_dtnt,0);
                END IF;
            END LOOP;

            OPEN c_Ma_Thue;
            FETCH c_Ma_Thue INTO v_Ma_Thue;
            CLOSE c_Ma_Thue;

            --Xu ly du lieu voi bang phat sinh
            INSERT INTO qlt_psinh_tkhai (id,
                                         tkh_id,
                                         tkh_ltd,
                                         ccg_ma_cap,
                                         ccg_ma_chuong,
                                         lkn_ma_loai,
                                         lkn_ma_khoan,
                                         tmt_ma_muc,
                                         tmt_ma_tmuc,
                                         tmt_ma_thue,
                                         thue_psinh,
                                         can_cu_tt)
            VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    p_Record_Dtnt.ma_cap,
                    p_Record_Dtnt.ma_chuong,
                    p_Record_Dtnt.ma_loai,
                    p_Record_Dtnt.ma_khoan,
                    '002',
                    '02',
                    v_Ma_Thue,
                    v_Thue_PSinh,
                    0);
        END IF;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        /*Ghi so ho so nhan*/
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '26',-- To khai TNDN Quy
                         '02',-- To khai nop thue
                         '105');
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_TNDN_Quy'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Khainhg
Ngay lap: 19/04/2006
Noi dung: Thuc hien do du lieu to khai thue tai nguyen vao trong CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai

********************************************************************************/
    PROCEDURE Prc_TKhai_TNguyen(p_Record_Of_Header Record_Hdr,
                                    p_Record_Dtnt Record_Dtnt,
                                    p_TKhai_Exits_Id NUMBER,
                                    p_Tthai_Tkhai VARCHAR2) IS

	    TYPE Record_Dtl_TNguyen IS RECORD(btn_id NUMBER(10,0),
                                          dvt_don_vi_tinh VARCHAR2(10),
                                          san_luong NUMBER(20,3),
                                          gia_don_vi NUMBER(20,2),
                                          gia_tt_don_vi NUMBER(20,2),
                                          tsuat_dtnt NUMBER(5,2),
                                          thue_phai_nop_dtnt NUMBER(20,2),
                                          tsuat_cqt NUMBER(5,2),
                                          thue_phai_nop_cqt NUMBER(20,2),
                                          thue_psinh_tky_dtnt NUMBER(20,2),
                                          thue_psinh_tky_cqt NUMBER(20,2),
                                          thue_mien_giam_dtnt NUMBER(20,2),
                                          thue_mien_giam_cqt NUMBER(20,2),
                                          tong_tri_gia_tt NUMBER(20,2),
                                          ke_khai_sai VARCHAR2(1));

        TYPE Record_Dtl_PSINH IS RECORD(ma_muc VARCHAR2(3),
                                        ma_tmuc VARCHAR2(2),
                                        thue_psinh NUMBER(20,2));

        TYPE Array_Tkhai_Dtl IS TABLE OF Record_Dtl_TNGUYEN INDEX BY BINARY_INTEGER;

        TYPE Array_Psinh_Dtl IS TABLE OF Record_Dtl_PSINH INDEX BY BINARY_INTEGER;

        v_Array_Tkhai_Dtl Array_Tkhai_Dtl;
        v_Array_Tkhai_Dtl_Tong Array_Tkhai_Dtl;

        v_Array_Psinh_Dtl Array_Psinh_Dtl;
        /* Lay chi tiet ke khai thue TNGUYEN */
        CURSOR c_TKhai_TNguyen IS
            SELECT *
            FROM rcv_v_tkhai_tnguyen
            WHERE (hdr_id = p_Record_Of_Header.id)
              AND (btn_id IS NOT NULL);

        /* Lay so phat sinh tren to khai TNGUYEN da ton tai */
        CURSOR c_Psinh_TNguyen_Exist(p_ma_muc VARCHAR2
                                   , p_ma_tmuc VARCHAR2) IS
            SELECT * FROM qlt_psinh_tkhai
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0)
              AND (tmt_ma_muc = p_ma_muc)
              AND (tmt_ma_tmuc = p_ma_tmuc);

        /* Lay dia diem khai thac, chi tieu tong */
        CURSOR c_Ctieu_Tong IS
            SELECT ddiem_kthac
                  ,thue_psinh_tky_dtnt
                  ,thue_mien_giam_dtnt
                  ,thue_phai_nop_dtnt
            FROM rcv_v_tkhai_tnguyen
            WHERE (hdr_id = p_Record_Of_Header.id);

        /* Tong hop so phat sinh */
        CURSOR c_Psinh_TNguyen IS
            SELECT  bthue.ma_muc,
                    bthue.ma_tmuc,
                    SUM(thue_psinh_tky_dtnt) thue_psinh
            FROM rcv_v_tkhai_tnguyen tkhai
               , rcv_map_ctieu_bthue bthue
               , qlt_dm_bthue_tnguyen dm
            WHERE (tkhai.btn_id IS NOT NULL)
              AND (tkhai.hdr_id = p_Record_Of_Header.id)
              AND (tkhai.btn_id = dm.id)
              AND (dm.ma = bthue.ma_ctieu)
              AND (bthue.loai_tkhai = '06')
              AND (dm.ngay_hl = (SELECT MAX(ngay_hl)
                                 FROM qlt_dm_bthue_tnguyen
                                 WHERE (ngay_hl < p_Record_Of_Header.kylb_tu_ngay)))
            GROUP BY bthue.ma_muc, bthue.ma_tmuc;

        /* Lay ma loai thue */
        CURSOR c_Loai_Thue IS
            SELECT lte_ma_thue
            FROM qlt_dm_tkhai
            WHERE (ma = p_Record_Of_Header.loai_tkhai);

        vc_Loai_Thue c_Loai_Thue%ROWTYPE;
        vc_Ctieu_Tong c_Ctieu_Tong%ROWTYPE;
        vc_Psinh_TNguyen_Exist c_Psinh_TNguyen_Exist%ROWTYPE;
        v_HanNop DATE;
        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_Ma_Thue VARCHAR(2);
        v_Thue_PSinh NUMBER(20,2);
        v_Check BOOLEAN := TRUE;
        v_Count NUMBER(10,0) := 0;
        v_Index NUMBER(10,0) :=0;
        v_Sai_SoHoc VARCHAR2(1) := NULL;
        v_Thue_Psinh_Tky_Cqt NUMBER;
        v_Thue_Phai_Nop_Cqt NUMBER;
        v_tmuc_exist VARCHAR2(100) := NULL;

    BEGIN
        --Tinh so phat sinh
        OPEN c_Loai_Thue;
        FETCH c_Loai_Thue INTO vc_Loai_Thue;
        IF (c_Loai_Thue%FOUND) THEN
            v_Ma_Thue := vc_Loai_Thue.lte_ma_thue;
        END IF;
        CLOSE c_Loai_Thue;

        --lay dia diem khai thac, ctieu tong
        OPEN c_Ctieu_Tong;
        FETCH c_Ctieu_Tong INTO vc_Ctieu_Tong;
        CLOSE c_Ctieu_Tong;

        /* Luu cac gia tri cua to khai TNGUYEN truc tiep ra mang */
        FOR vc_TKhai_TNguyen IN c_TKhai_TNguyen LOOP
            v_Count := v_Count + 1;
            v_Array_Tkhai_Dtl(v_Count).btn_id := vc_TKhai_TNguyen.btn_id;
            v_Array_Tkhai_Dtl(v_Count).dvt_don_vi_tinh := vc_TKhai_TNguyen.don_vi_tinh;
            v_Array_Tkhai_Dtl(v_Count).san_luong := vc_TKhai_TNguyen.san_luong;
            v_Array_Tkhai_Dtl(v_Count).gia_don_vi := vc_TKhai_TNguyen.gia_don_vi;
            v_Array_Tkhai_Dtl(v_Count).gia_tt_don_vi := vc_TKhai_TNguyen.gia_tt_don_vi;
            v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt := vc_TKhai_TNguyen.tsuat_dtnt;
            v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_dtnt := vc_TKhai_TNguyen.thue_phai_nop_dtnt;
            v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_dtnt := vc_TKhai_TNguyen.thue_psinh_tky_dtnt;
            v_Array_Tkhai_Dtl(v_Count).thue_mien_giam_dtnt := NVL(vc_TKhai_TNguyen.thue_mien_giam_dtnt,0);
            /* Tong tri gia tinh thue = san luong * gia don vi */
            v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt := NVL(v_Array_Tkhai_Dtl(v_Count).san_luong,0)
                                                        * NVL(v_Array_Tkhai_Dtl(v_Count).gia_don_vi,0);
            /*
                - Tinh so co quan thue.
                - Neu Gia tinh thue don vi tai nguyen = 0
                    -> Thue psinh trong ky = San luong
                                           * muc thue an dinh tren 1 dv tai nguyen
                  Else Thue psinh trong ky = San luong
                                           * Gia tinh thue don vi * thue suat/100
            */
            IF NVL(v_Array_Tkhai_Dtl(v_Count).gia_don_vi,0) =0 THEN
                v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_cqt := v_Array_Tkhai_Dtl(v_Count).san_luong
                                                               * v_Array_Tkhai_Dtl(v_Count).gia_tt_don_vi;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_cqt := Round(v_Array_Tkhai_Dtl(v_Count).san_luong
                                                               * v_Array_Tkhai_Dtl(v_Count).gia_don_vi
                                                               * v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt/100);
            END IF;
            v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_cqt := v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_cqt
                                                          - v_Array_Tkhai_Dtl(v_Count).thue_mien_giam_dtnt;
            /*
                - Kiem tra so dtnt va so cqt
                - Neu thue psinh tky Dtnt <> thue psinh tky cqt
                    or thue pnop dtnt <> thue pnop cqt (kiem tra theo nguong) -> ke khai sai = 'Y'
            */
            IF (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_dtnt,
                               v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_cqt,
                               p_Record_Of_Header.kylb_tu_ngay))
                AND (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_dtnt,
                                    v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_cqt,
                                    p_Record_Of_Header.kylb_tu_ngay)) THEN
                v_Array_Tkhai_Dtl(v_Count).ke_khai_sai := NULL;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).ke_khai_sai := 'Y';
                v_Sai_SoHoc := 'Y';
            END IF;
            /* Tinh chi tieu tong */
            v_Thue_Psinh_Tky_Cqt := NVL(v_Thue_Psinh_Tky_Cqt,0)
                                  + NVL(v_Array_Tkhai_Dtl(v_Count).thue_psinh_tky_cqt,0);
            v_Thue_Phai_Nop_Cqt := NVL(v_Thue_Phai_Nop_Cqt,0)
                                 + NVL(v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_cqt,0);
        END LOOP;

        /* Tinh chi tieu tong DTNT */
        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_cqt := v_Thue_Psinh_Tky_Cqt;
        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_cqt := v_Thue_Phai_Nop_Cqt;
        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_dtnt := NVL(vc_Ctieu_Tong.thue_psinh_tky_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).thue_mien_giam_dtnt := NVL(vc_Ctieu_Tong.thue_mien_giam_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_dtnt := NVL(vc_Ctieu_Tong.thue_phai_nop_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).btn_id := 68;
        /*
            - Kiem tra so dtnt va so cqt
            - Neu thue psinh tky Dtnt <> thue psinh tky cqt
                or thue pnop dtnt <> thue pnop cqt (kiem tra theo nguong) -> ke khai sai = 'Y'
        */
        IF (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_dtnt,
                           v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_cqt,
                           p_Record_Of_Header.kylb_tu_ngay))
            AND (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_dtnt,
                                v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_cqt,
                                p_Record_Of_Header.kylb_tu_ngay)) THEN
            v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai := NULL;
        ELSE
            v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai := 'Y';
            v_Sai_SoHoc := 'Y';
        END IF;

        /* Tinh so phat sinh */
        FOR vc_Psinh_TNguyen IN c_Psinh_TNguyen LOOP
            v_Index := v_Index +1;
            v_Array_Psinh_Dtl(v_Index).ma_muc := NVL(vc_Psinh_TNguyen.ma_muc,'012');
            v_Array_Psinh_Dtl(v_Index).ma_tmuc := NVL(vc_Psinh_TNguyen.ma_tmuc,'01');
            v_Array_Psinh_Dtl(v_Index).thue_psinh := NVL(vc_Psinh_TNguyen.thue_psinh,0);
        END LOOP;
        /*Neu da ton tai to khai trong ky ke khai*/
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            Qlt_Pck_Gdich.Prc_Lay_Thamso(p_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);
            --Xoa du lieu trong bang detail voi dieu kien ltd =0
            DELETE FROM qlt_tkhai_tnguyen
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

            /*Thong tin Header*/
            UPDATE qlt_tkhai_hdr
               SET co_loi = v_Sai_SoHoc,
                   co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = '4', --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat,
                   ddiem_kthac = vc_Ctieu_Tong.ddiem_kthac
              WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);
            /*Thong tin Detail*/
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_tnguyen (id,
                                               tkh_id,
                                               tkh_ltd,
                                               btn_id,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_don_vi,
                                               gia_tt_don_vi,
                                               tsuat_dtnt,
                                               thue_phai_nop_dtnt,
                                               tsuat_cqt,
                                               thue_phai_nop_cqt,
                                               thue_psinh_tky_dtnt,
                                               thue_psinh_tky_cqt,
                                               thue_mien_giam_dtnt,
                                               thue_mien_giam_cqt,
                                               tong_tri_gia_tt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        p_TKhai_Exits_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btn_id,
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).san_luong,
                        v_Array_Tkhai_Dtl(i).gia_don_vi,
                        v_Array_Tkhai_Dtl(i).gia_tt_don_vi,
                        v_Array_Tkhai_Dtl(i).tsuat_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt,
                        v_Array_Tkhai_Dtl(i).tsuat_cqt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_tky_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_tky_cqt,
                        v_Array_Tkhai_Dtl(i).thue_mien_giam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_mien_giam_cqt,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai);
            END LOOP;
            /* Insert chi tieu tong so */
                INSERT INTO qlt_tkhai_tnguyen (id,
                                               tkh_id,
                                               tkh_ltd,
                                               btn_id,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_don_vi,
                                               gia_tt_don_vi,
                                               tsuat_dtnt,
                                               thue_phai_nop_dtnt,
                                               tsuat_cqt,
                                               thue_phai_nop_cqt,
                                               thue_psinh_tky_dtnt,
                                               thue_psinh_tky_cqt,
                                               thue_mien_giam_dtnt,
                                               thue_mien_giam_cqt,
                                               tong_tri_gia_tt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        p_TKhai_Exits_Id,
                        0,
                        v_Array_Tkhai_Dtl_Tong(1).btn_id,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_dtnt,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_mien_giam_dtnt,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai);
            /*Thong tin bang phat sinh*/
            FOR i iN 1..v_Index LOOP
                --Danh sach tieu muc co tren to khai trong mode sua
                v_tmuc_exist := v_tmuc_exist ||','||v_Array_Psinh_Dtl(i).ma_tmuc;
                --Lay so phat sinh tren to khai truoc
                OPEN c_Psinh_TNguyen_Exist(v_Array_Psinh_Dtl(i).ma_muc,
                                           v_Array_Psinh_Dtl(i).ma_tmuc);
                FETCH c_Psinh_TNguyen_Exist INTO vc_Psinh_TNguyen_Exist;
                IF c_Psinh_TNguyen_Exist%FOUND THEN
                --Neu muc phat sinh da ton tai
                    IF vc_Psinh_TNguyen_Exist.thue_psinh <> v_Array_Psinh_Dtl(i).thue_psinh THEN
                        UPDATE qlt_psinh_tkhai
                        SET thue_psinh = v_Array_Psinh_Dtl(i).thue_psinh
                        WHERE (tkh_id = p_TKhai_Exits_Id)
                          AND (tkh_ltd = 0)
                          AND (tmt_ma_muc = vc_Psinh_TNguyen_Exist.tmt_ma_muc)
                          AND (tmt_ma_tmuc = vc_Psinh_TNguyen_Exist.tmt_ma_tmuc);
                    END IF;
                /* Neu muc chua ton tai */
                ELSE
                        INSERT INTO qlt_psinh_tkhai (id,
                                                     tkh_id,
                                                     tkh_ltd,
                                                     ccg_ma_cap,
                                                     ccg_ma_chuong,
                                                     lkn_ma_loai,
                                                     lkn_ma_khoan,
                                                     tmt_ma_muc,
                                                     tmt_ma_tmuc,
                                                     tmt_ma_thue,
                                                     thue_psinh)
                        VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                                p_TKhai_Exits_Id,
                                0,
                                p_Record_Dtnt.ma_cap,
                                p_Record_Dtnt.ma_chuong,
                                p_Record_Dtnt.ma_loai,
                                p_Record_Dtnt.ma_khoan,
                                v_Array_Psinh_Dtl(i).ma_muc,
                                v_Array_Psinh_Dtl(i).ma_tmuc,
                                v_Ma_Thue,
                                v_Array_Psinh_Dtl(i).thue_psinh);
                END IF;
                CLOSE c_Psinh_TNguyen_Exist;
            END LOOP;

            --Xoa nhung muc khong ton tai
            DELETE FROM qlt_psinh_tkhai
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0)
              AND (INSTR(v_tmuc_exist,tmt_ma_tmuc) = 0);
        ELSE --To khai chua ton tai
            --Neu co an dinh
            IF (Fnc_AnDinh_Exits(p_Record_Of_Header.tin,
                                 p_Record_Of_Header.loai_tkhai,
                                 p_Record_Of_Header.kykk_tu_ngay,
                                 p_Record_Of_Header.kykk_den_ngay)) THEN
                v_Tthai_Tkhai := '3'; --To khai nop cham
            ELSE --Neu chua co an dinh
                v_HanNop := Fnc_Han_Nop(p_Record_Of_Header.loai_tkhai,
                                        'TK',
                                        p_Record_Of_Header.kykk_den_ngay);
                IF (p_Record_Of_Header.ngay_nop <= v_HanNop) THEN
                    --To khai dung han
                    v_Tthai_Tkhai := '1'; --To khai chinh thuc
                ELSE
                    --To khai khong dung han
                    v_Tthai_Tkhai := '3'; --To khai nop cham
                END IF;
            END IF;

            --Thuc hien xu ly to khai
            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai,
                                         p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Xu ly voi du lieu to khai Header
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
                                      co_loi,
                                      co_loi_ddanh,
                                      ghi_chu_loi,
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat,
                                      ddiem_kthac)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   v_Sai_SoHoc,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   NULL,  --p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat,
                   vc_Ctieu_Tong.ddiem_kthac);

            --Xu ly voi du lieu to khai Detail
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_tnguyen (id,
                                               tkh_id,
                                               tkh_ltd,
                                               btn_id,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_don_vi,
                                               gia_tt_don_vi,
                                               tsuat_dtnt,
                                               thue_phai_nop_dtnt,
                                               tsuat_cqt,
                                               thue_phai_nop_cqt,
                                               thue_psinh_tky_dtnt,
                                               thue_psinh_tky_cqt,
                                               thue_mien_giam_dtnt,
                                               thue_mien_giam_cqt,
                                               tong_tri_gia_tt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btn_id,
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).san_luong,
                        v_Array_Tkhai_Dtl(i).gia_don_vi,
                        v_Array_Tkhai_Dtl(i).gia_tt_don_vi,
                        v_Array_Tkhai_Dtl(i).tsuat_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt,
                        v_Array_Tkhai_Dtl(i).tsuat_cqt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_tky_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_tky_cqt,
                        v_Array_Tkhai_Dtl(i).thue_mien_giam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_mien_giam_cqt,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai);
            END LOOP;
            /* Insert chi tieu tong so */
                INSERT INTO qlt_tkhai_tnguyen (id,
                                               tkh_id,
                                               tkh_ltd,
                                               btn_id,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_don_vi,
                                               gia_tt_don_vi,
                                               tsuat_dtnt,
                                               thue_phai_nop_dtnt,
                                               tsuat_cqt,
                                               thue_phai_nop_cqt,
                                               thue_psinh_tky_dtnt,
                                               thue_psinh_tky_cqt,
                                               thue_mien_giam_dtnt,
                                               thue_mien_giam_cqt,
                                               tong_tri_gia_tt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Tkhai_Dtl_Tong(1).btn_id,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_dtnt,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_tky_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_mien_giam_dtnt,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai);
            --Insert du lieu vao bang phat sinh
            FOR i IN 1..v_Index LOOP
                    INSERT INTO qlt_psinh_tkhai (id,
                                                 tkh_id,
                                                 tkh_ltd,
                                                 ccg_ma_cap,
                                                 ccg_ma_chuong,
                                                 lkn_ma_loai,
                                                 lkn_ma_khoan,
                                                 tmt_ma_muc,
                                                 tmt_ma_tmuc,
                                                 tmt_ma_thue,
                                                 thue_psinh)
                    VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                            v_Hdr_Id,
                            0,
                            p_Record_Dtnt.ma_cap,
                            p_Record_Dtnt.ma_chuong,
                            p_Record_Dtnt.ma_loai,
                            p_Record_Dtnt.ma_khoan,
                            v_Array_Psinh_Dtl(i).ma_muc,
                            v_Array_Psinh_Dtl(i).ma_tmuc,
                            '04',--v_Ma_Thue,
                            v_Array_Psinh_Dtl(i).thue_psinh);
            END LOOP;
        END IF;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        /*Ghi thong tin ho so nhan*/
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '24', --To khai Tnguyen
                         '02', --To khai thue
                         '102');
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_Tnguyen'
                                        , FALSE
                                        , NULL);

    END;
/*******************************************************************************
Nguoi lap: Khainhg
Ngay lap: 26/04/2006
Noi dung: Thuc hien do du lieu to khai quyet toan thue tai nguyen vao trong CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai

*******************************************************************************/
    PROCEDURE Prc_TKhai_TTDB(p_Record_Of_Header Record_Hdr,
                                p_Record_Dtnt Record_Dtnt,
                                p_TKhai_Exits_Id NUMBER,
                                p_Tthai_Tkhai VARCHAR2) IS

	    TYPE Record_Dtl_TTDB IS RECORD(btt_id NUMBER(10,0),
                                       ma_ctieu VARCHAR2(6),
                                       so_tt NUMBER(10,0),
                                       dvt_don_vi_tinh VARCHAR2(10),
                                       so_luong NUMBER(20,3),
                                       tong_tri_gia_ban NUMBER(20,2),
                                       tong_tri_gia_tt_dtnt NUMBER(20,2),
                                       thue_duoc_ktru NUMBER(20,2),
                                       tsuat_dtnt NUMBER(5,2),
                                       thue_phai_nop_dtnt NUMBER(20,2),
                                       tong_tri_gia_tt_cqt NUMBER(20,2),
                                       thue_phai_nop_cqt NUMBER(20,2),
                                       dchinh_tang_giam NUMBER(20,2),
                                       thue_pnop_tky_dtnt NUMBER(20,2),
                                       thue_pnop_tky_cqt NUMBER(20,2),
                                       ke_khai_sai CHAR(1),
                                       giatri_baobi NUMBER(20,2),
                                       kieu_thue_pnop_tky_cqt VARCHAR2(1),
                                       kieu_tong_tri_gia_tt_cqt VARCHAR2(1));

        TYPE Record_Dtl_PSINH IS RECORD(ma_muc VARCHAR2(3),
                                        ma_tmuc VARCHAR2(2),
                                        thue_psinh NUMBER(20,2));

        TYPE Record_Pluc_01C IS RECORD(id NUMBER(10,0),
                                      tkh_id NUMBER(10,0),
                                      tkh_ltd NUMBER(10,0),
                                      btt_id NUMBER(10,0),
                                      kykk_tu_ngay DATE,
                                      kykk_den_ngay DATE,
                                      ctg_id NUMBER(10,0),
                                      so_kkhai NUMBER(20,2),
                                      so_dchinh NUMBER(20,2),
                                      so_clech_dtnt NUMBER(20,2),
                                      so_clech_cqt NUMBER(20,2),
                                      ly_do_dchinh VARCHAR2(250),
                                      ke_khai_sai VARCHAR2(1));

        TYPE Array_Tkhai_Dtl IS TABLE OF Record_Dtl_TTDB INDEX BY BINARY_INTEGER;

        TYPE Array_Psinh_Dtl IS TABLE OF Record_Dtl_PSINH INDEX BY BINARY_INTEGER;

        TYPE Array_Pluc_01C IS TABLE OF Record_Pluc_01C INDEX BY BINARY_INTEGER;

        v_Array_Tkhai_Dtl Array_Tkhai_Dtl;

        v_Array_Psinh_Dtl Array_Psinh_Dtl;

        v_Array_Pluc_01C Array_Pluc_01C;
        /* Lay chi tiet ke khai thue TTDB */
        CURSOR c_TKhai_TTDB(p_Btt_Id VARCHAR2 DEFAULT NULL) IS
            SELECT *
            FROM rcv_v_tkhai_ttdb
            WHERE (hdr_id = p_Record_Of_Header.id)
            AND ((p_Btt_Id IS NULL AND btt_id <> 154) OR btt_id = p_Btt_Id)
            AND btt_id IS NOT NULL
            ORDER BY id_tt;
        /* Lay ma chi tieu */
        CURSOR c_Ctieu_Tttb(p_Id NUMBER) IS
            SELECT ma ma_ctieu, giatri_baobi
            FROM qlt_dm_bthue_ttdb
            WHERE (id = p_Id)
            AND (ngay_hl = (SELECT MAX(ngay_hl)
                           FROM qlt_dm_bthue_ttdb
                           WHERE ngay_hl <= p_Record_Of_Header.kylb_tu_ngay));
        /* Lay chi tiet phu luc */
        CURSOR c_Pluc_Ttdb_01C IS
            SELECT *
            FROM rcv_v_pluc_ttdb_01c
            WHERE (hdr_id = p_Record_Of_Header.id)
            AND (ctg_id IS NULL OR ctg_id <> 13) --bo chi tieu don vi tinh
            AND ((btt_id IS NOT NULL)
              OR (ctg_id IS NOT NULL AND (NVL(so_kkhai,0) <> 0
                                       OR NVL(so_dchinh,0) <> 0
                                       OR NVL(so_clech_dtnt,0) <> 0)))
            ORDER BY id_tt;
        /* Lay so phat sinh tren to khai TNGUYEN da ton tai */
        CURSOR c_Psinh_TNguyen_Exist(p_ma_muc VARCHAR2
                                   , p_ma_tmuc VARCHAR2) IS
            SELECT * FROM qlt_psinh_tkhai
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0)
              AND (tmt_ma_muc = p_ma_muc)
              AND (tmt_ma_tmuc = p_ma_tmuc);
        /* Tong hop so phat sinh */
        CURSOR c_Psinh_TNguyen IS
            SELECT  bthue.ma_muc,
                    bthue.ma_tmuc,
                    SUM(thue_pnop_tky_dtnt) thue_psinh
            FROM rcv_v_tkhai_ttdb tkhai
               , rcv_map_ctieu_bthue bthue
               , qlt_dm_bthue_ttdb dm
            WHERE (tkhai.btt_id IS NOT NULL)
              AND (tkhai.hdr_id = p_Record_Of_Header.id)
              AND (tkhai.btt_id = dm.id)
              AND (dm.ma = bthue.ma_ctieu)
              AND (bthue.loai_tkhai = '05')
              AND (dm.ngay_hl = (SELECT MAX(ngay_hl)
                                 FROM qlt_dm_bthue_ttdb
                                 WHERE (ngay_hl < p_Record_Of_Header.kylb_tu_ngay)))
            GROUP BY bthue.ma_muc, bthue.ma_tmuc;

        /* Lay ma loai thue */
        CURSOR c_Loai_Thue IS
            SELECT lte_ma_thue
            FROM qlt_dm_tkhai
            WHERE (ma = p_Record_Of_Header.loai_tkhai);

        vc_Loai_Thue c_Loai_Thue%ROWTYPE;
        vc_Ctieu_Tttb c_Ctieu_Tttb%ROWTYPE;
        vc_TKhai_TTDB c_TKhai_TTDB%ROWTYPE;
        vc_Psinh_TNguyen_Exist c_Psinh_TNguyen_Exist%ROWTYPE;
        v_HanNop DATE;
        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_Btt_Id NUMBER(10);
        v_Ma_Thue VARCHAR(2);
        v_Check BOOLEAN := TRUE;
        v_Count NUMBER(10) := 0;
        v_Index NUMBER(10) :=0;
        v_Count_Pluc NUMBER(10) := 0;
        --Bien su dung cho ghi so nhan ho so
        v_Exist_Pluc BOOLEAN := FALSE;
        v_Ds_Pluc VARCHAR2(10);

        v_Sai_SoHoc VARCHAR2(1) := NULL;
    	v_Tong_Phai_nop_Cqt NUMBER(20,2):= 0;
    	v_Tong_Phai_nop_Tky_Cqt NUMBER(20,2):= 0;
    	v_Tong_Phai_nop_I_Cqt NUMBER(20,2):= 0;
    	v_Tong_Phai_nop_Tky_I_Cqt NUMBER(20,2):= 0;
    	v_Tong_Phai_nop_II_Cqt NUMBER(20,2):= 0;
    	v_Tong_Phai_nop_Tky_II_Cqt NUMBER(20,2):= 0;
        v_Kykk_Tu_Ngay DATE;
        v_tmuc_exist VARCHAR2(100);
    BEGIN
        --Lay loai thue
        OPEN c_Loai_Thue;
        FETCH c_Loai_Thue INTO vc_Loai_Thue;
        IF (c_Loai_Thue%FOUND) THEN
            v_Ma_Thue := vc_Loai_Thue.lte_ma_thue;
        END IF;
        CLOSE c_Loai_Thue;
        /* Luu cac gia tri cua to khai TTDB truc tiep ra mang */
        FOR vc_TKhai_TTDB IN c_TKhai_TTDB LOOP
            v_Count := v_Count + 1;
            --Lay ma ctieu
            OPEN c_Ctieu_Tttb(vc_TKhai_TTDB.btt_id);
            FETCH c_Ctieu_Tttb INTO vc_Ctieu_Tttb;
            CLOSE c_Ctieu_Tttb;
            /* Gan kieu du lieu */
            v_Array_Tkhai_Dtl(v_Count).kieu_thue_pnop_tky_cqt := vc_TKhai_TTDB.kieu_thue_pnop_tky_dtnt;
            v_Array_Tkhai_Dtl(v_Count).kieu_tong_tri_gia_tt_cqt := vc_TKhai_TTDB.kieu_tong_tri_gia_tt_dtnt;
            /* Gan gia tri */
            v_Array_Tkhai_Dtl(v_Count).so_tt := v_Count +1;
            v_Array_Tkhai_Dtl(v_Count).btt_id := vc_TKhai_TTDB.btt_id;

            v_Array_Tkhai_Dtl(v_Count).ma_ctieu := vc_Ctieu_Tttb.ma_ctieu;
            v_Array_Tkhai_Dtl(v_Count).dvt_don_vi_tinh := vc_TKhai_TTDB.dvt_don_vi_tinh;

            v_Array_Tkhai_Dtl(v_Count).so_luong := vc_TKhai_TTDB.so_luong;
            v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_ban := vc_TKhai_TTDB.tong_tri_gia_ban;

            v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt := REPLACE(NVL(vc_TKhai_TTDB.tsuat_dtnt,0),',','.');
            v_Array_Tkhai_Dtl(v_Count).thue_duoc_ktru := vc_TKhai_TTDB.thue_duoc_ktru;

            v_Array_Tkhai_Dtl(v_Count).dchinh_tang_giam := vc_TKhai_TTDB.dchinh_tang_giam;
            v_Array_Tkhai_Dtl(v_Count).giatri_baobi := vc_Ctieu_Tttb.giatri_baobi;
            --Tinh so dtnt
            v_Array_Tkhai_Dtl(v_Count).thue_pnop_tky_dtnt := vc_TKhai_TTDB.thue_pnop_tky_dtnt;
            v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_dtnt := vc_TKhai_TTDB.tong_tri_gia_tt_dtnt;
            /* Tinh so CQT */
            --Tong tri gia tinh thue
            IF v_Array_Tkhai_Dtl(v_Count).kieu_tong_tri_gia_tt_cqt IS NULL THEN
                v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt := NULL;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt := (NVL(v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_ban,0)
                                                                - NVL(v_Array_Tkhai_Dtl(v_Count).giatri_baobi,0)
                                                                * NVL(v_Array_Tkhai_Dtl(v_Count).so_luong,0))
                                                                / (1 + NVL(v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt,0)/100);
            END IF;

            IF v_Array_Tkhai_Dtl(v_Count).ma_ctieu IN ('I','II','C')
                AND v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_dtnt IS NOT NULL  THEN
                v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt := v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_dtnt;
            END IF;
            --So phai nop trong ky
            IF v_Array_Tkhai_Dtl(v_Count).kieu_thue_pnop_tky_cqt IS NULL THEN
                v_Array_Tkhai_Dtl(v_Count).thue_pnop_tky_cqt := NULL;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).thue_pnop_tky_cqt := NVL(v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt,0)
                                                              * (NVL(v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt,0)/100)
                                                              - NVL(v_Array_Tkhai_Dtl(v_Count).thue_duoc_ktru,0)
                                                              + NVL(v_Array_Tkhai_Dtl(v_Count).dchinh_tang_giam,0);
            END IF;
            /*
                - So phai nop CQT
                - Thue_phai_nop_cqt := tong_tri_gia_tt_cqt * tsuat_dtnt/100
                - Thue_phai_nop_dtnt := tong_tri_gia_tt_cqt * tsuat_dtnt/100
            */
            v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_cqt := NVL(v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt,0)
                                                          * (NVL(v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt,0)/100);
            v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_dtnt := NVL(v_Array_Tkhai_Dtl(v_Count).tong_tri_gia_tt_cqt,0)
                                                          * (NVL(v_Array_Tkhai_Dtl(v_Count).tsuat_dtnt,0)/100);

            /*
                Neu ma_ctieu Not in (I,II,III) <- Chi tieu tong
                - Tinh tong phai nop trong ky (Cac chi tieu order by theo thu tu)
            */
            IF v_Array_Tkhai_Dtl(v_Count).ma_ctieu NOT IN ('I','II','III') THEN
                v_Tong_Phai_Nop_Tky_Cqt  := v_Tong_Phai_Nop_Tky_Cqt
                                          + NVL(v_Array_Tkhai_Dtl(v_Count).thue_pnop_tky_cqt,0);
                v_Tong_Phai_Nop_Cqt  := v_Tong_Phai_Nop_Cqt
                                      + NVL(v_Array_Tkhai_Dtl(v_Count).thue_phai_nop_cqt,0);
            END IF;
            /*
                Neu ma_ctieu = 'II' -> gan gia tri tinh tong den dong co ma = 'II' cho tong I
            */
            IF v_Array_Tkhai_Dtl(v_Count).ma_ctieu = 'II' THEN
    			v_Tong_Phai_Nop_I_Cqt := v_tong_phai_nop_Cqt;
				v_Tong_Phai_Nop_Tky_I_Cqt := v_tong_phai_nop_Tky_Cqt;
				v_Tong_Phai_Nop_Cqt := 0;
				v_Tong_Phai_Nop_Tky_Cqt := 0;
            END IF;
            /*
                Neu ma_ctieu = 'III' -> gan gia tri tinh tong den dong co ma = 'III' cho tong II
            */
            IF v_Array_Tkhai_Dtl(v_Count).ma_ctieu = 'III' THEN
    			v_Tong_phai_nop_II_Cqt  := v_tong_phai_nop_Cqt;
				v_Tong_phai_nop_Tky_II_Cqt  := v_tong_phai_nop_Tky_Cqt;
				v_Tong_phai_nop_Cqt  := 0;
				v_Tong_phai_nop_Tky_Cqt  := 0;
            END IF;
        END LOOP;
        /* Tinh gia tri cho cac chi tieu tong */
        FOR i IN 1..v_Count LOOP
            --Gan gia tri cac chi tieu tong vao mang
            IF v_Array_Tkhai_Dtl(i).ma_ctieu = 'I' THEN
                v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt := v_Tong_Phai_Nop_Tky_I_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt := v_Tong_Phai_Nop_I_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt := v_Tong_Phai_Nop_I_Cqt;

            ELSIF v_Array_Tkhai_Dtl(i).ma_ctieu = 'II' THEN
                v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt := v_Tong_Phai_Nop_Tky_II_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt := v_Tong_Phai_Nop_II_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt := v_Tong_Phai_Nop_II_Cqt;

            ELSIF v_Array_Tkhai_Dtl(i).ma_ctieu = 'C' THEN
                v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt := v_Tong_Phai_Nop_Tky_II_Cqt
                                                         + v_Tong_Phai_Nop_Tky_I_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt := v_Tong_Phai_Nop_II_Cqt
                                                        + v_Tong_Phai_Nop_I_Cqt;
                v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt := v_Tong_Phai_Nop_II_Cqt
                                                        + v_Tong_Phai_Nop_I_Cqt;
            END IF;
        END LOOP;
        /* Kiem tra so CQT va so DTNT*/
        FOR i IN 1..v_Count LOOP
            IF v_Array_Tkhai_Dtl(i).ma_ctieu NOT IN ('B','III') THEN
    			IF  Fnc_Ktra_Ctieu(ROUND(v_Array_Tkhai_Dtl(i).thue_pnop_tky_dtnt)
                                 , ROUND(v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt)
                                 , p_Record_Of_Header.kylb_tu_ngay)
                    AND Fnc_Ktra_Ctieu(ROUND(v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_dtnt)
                                     , ROUND(v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_cqt)
                                     , p_Record_Of_Header.kylb_tu_ngay) THEN
                    v_Array_Tkhai_Dtl(i).ke_khai_sai := NULL;
		      	ELSE
                    v_Array_Tkhai_Dtl(i).ke_khai_sai := 'Y';
    		      	v_Sai_SoHoc := 'Y';
		      END IF;
            END IF;
        END LOOP;
        /* Luu gia tri cua phu luc ra mang */
        FOR v_Pluc_Ttdb_01C IN c_Pluc_Ttdb_01C LOOP
            v_Exist_Pluc := TRUE; --Co phu luc
            /* Gan chi tieu, Kykk*/
            IF v_Pluc_Ttdb_01C.btt_id IS NOT NULL THEN
                v_Btt_Id := v_Pluc_Ttdb_01C.btt_id;
                v_Kykk_Tu_Ngay := To_Date('01'||v_Pluc_Ttdb_01C.ky_ke_khai,'DD/MM/YYYY');
            END IF;

            IF v_Pluc_Ttdb_01C.btt_id IS NULL THEN
                v_Count_Pluc := v_Count_Pluc + 1;
                v_Array_Pluc_01C(v_Count_Pluc).btt_id := v_Btt_Id;
                v_Array_Pluc_01C(v_Count_Pluc).kykk_tu_ngay := NVL(v_Kykk_Tu_Ngay,SYSDATE);
                v_Array_Pluc_01C(v_Count_Pluc).kykk_den_ngay := NVL(Last_Day(v_Kykk_Tu_Ngay),SYSDATE);
                v_Array_Pluc_01C(v_Count_Pluc).ctg_id := NVL(v_Pluc_Ttdb_01C.ctg_id,0);
                v_Array_Pluc_01C(v_Count_Pluc).so_kkhai := NVL(v_Pluc_Ttdb_01C.so_kkhai,0);
                v_Array_Pluc_01C(v_Count_Pluc).so_dchinh := NVL(v_Pluc_Ttdb_01C.so_dchinh,0);
                v_Array_Pluc_01C(v_Count_Pluc).so_clech_dtnt := NVL(v_Pluc_Ttdb_01C.so_clech_dtnt,0);
                v_Array_Pluc_01C(v_Count_Pluc).ly_do_dchinh := v_Pluc_Ttdb_01C.ly_do_dchinh;
                /*
                    - Tinh so CQT
                    - Neu so_clech_Dtnt is Null -> so_clech_dtnt = so_dchinh - so_kkhai
                      Else so_clech_cqt := so_clech_dtnt
                */
                IF v_Array_Pluc_01C(v_Count_Pluc).so_clech_dtnt IS NULL THEN
                    v_Array_Pluc_01C(v_Count_Pluc).so_clech_cqt := NVL(v_Array_Pluc_01C(v_Count_Pluc).so_dchinh,0)
                                                                 - NVL(v_Array_Pluc_01C(v_Count_Pluc).so_kkhai,0);
                ELSE
                    v_Array_Pluc_01C(v_Count_Pluc).so_clech_cqt := v_Array_Pluc_01C(v_Count_Pluc).so_clech_dtnt;
                END IF;
                /* Kiem tra so CQT va so DTNT */
                IF Fnc_Ktra_Ctieu(ROUND(v_Array_Pluc_01C(v_Count_Pluc).so_clech_dtnt)
                                , ROUND(v_Array_Pluc_01C(v_Count_Pluc).so_clech_cqt)
                                , p_Record_Of_Header.kylb_tu_ngay) THEN
                    v_Array_Pluc_01C(v_Count_Pluc).ke_khai_sai := NULL;
                ELSE
                    v_Array_Pluc_01C(v_Count_Pluc).ke_khai_sai := 'Y';
                    v_Sai_SoHoc := 'Y';
                END IF;
            END IF;
        END LOOP;
        /* Tinh so phat sinh */
        FOR vc_Psinh_TNguyen IN c_Psinh_TNguyen LOOP
            v_Index := v_Index +1;
            v_Array_Psinh_Dtl(v_Index).ma_muc := NVL(vc_Psinh_TNguyen.ma_muc,'015');
            v_Array_Psinh_Dtl(v_Index).ma_tmuc := NVL(vc_Psinh_TNguyen.ma_tmuc,'01');
            v_Array_Psinh_Dtl(v_Index).thue_psinh := NVL(vc_Psinh_TNguyen.thue_psinh,0);
        END LOOP;

        /*Neu da ton tai to khai trong ky ke khai*/
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            Qlt_Pck_Gdich.Prc_Lay_Thamso(p_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);
            --Xoa du lieu trong bang tkhai chi tiet voi dieu kien ltd =0
            DELETE qlt_tkhai_ttdb WHERE tkh_id = p_TKhai_Exits_Id
                                 AND tkh_ltd = 0;
            --Xoa du lieu phu luc
            DELETE qlt_pluc_ttdb_01c WHERE tkh_id = p_TKhai_Exits_Id
                                 AND tkh_ltd = 0;
            /*Thong tin Header*/
            UPDATE qlt_tkhai_hdr
               SET co_loi = v_Sai_SoHoc,
                   co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = '4', --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat
              WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);
            /*Thong tin Detail*/
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_ttdb (id,
                                            tkh_id,
                                            tkh_ltd,
                                            btt_id,
                                            dvt_don_vi_tinh,
                                            so_luong,
                                            tong_tri_gia_ban,
                                            tong_tri_gia_tt_dtnt,
                                            thue_duoc_ktru,
                                            tsuat_dtnt,
                                            thue_phai_nop_dtnt,
                                            tong_tri_gia_tt_cqt,
                                            thue_phai_nop_cqt,
                                            ke_khai_sai,
                                            dchinh_tang_giam,
                                            so_tt,
                                            thue_pnop_tky_dtnt,
                                            thue_pnop_tky_cqt)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        p_TKhai_Exits_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btt_id,
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).so_luong,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_ban,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_duoc_ktru,
                        v_Array_Tkhai_Dtl(i).tsuat_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_cqt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai,
                        v_Array_Tkhai_Dtl(i).dchinh_tang_giam,
                        v_Array_Tkhai_Dtl(i).so_tt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_tky_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt);
            END LOOP;
            /*Insert chitieu khong phat sinh tri gia tinh thue TTDB trong ky*/
            OPEN c_TKhai_TTDB(154);
            FETCH c_TKhai_TTDB INTO vc_TKhai_TTDB;
            CLOSE c_TKhai_TTDB;
            INSERT INTO qlt_tkhai_ttdb (id,
                            tkh_id,
                            tkh_ltd,
                            btt_id,
                            dvt_don_vi_tinh,
                            so_luong,
                            tong_tri_gia_ban,
                            tong_tri_gia_tt_dtnt,
                            thue_duoc_ktru,
                            tsuat_dtnt,
                            thue_phai_nop_dtnt,
                            tong_tri_gia_tt_cqt,
                            thue_phai_nop_cqt,
                            dchinh_tang_giam,
                            so_tt,
                            thue_pnop_tky_dtnt,
                            thue_pnop_tky_cqt)
            VALUES(qlt_xltk_dtl_seq.NEXTVAL,
                   p_TKhai_Exits_Id,
                   0,
                   154,
                   NULL,
                   DECODE(UPPER(vc_TKhai_TTDB.dvt_don_vi_tinh),'X',1,0),
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   1,
                   0,
                   0);
            /*Insert phu luc to khai*/
            FOR i IN 1..v_Count_Pluc LOOP
                INSERT INTO qlt_pluc_ttdb_01c(id,
                                              tkh_id,
                                              tkh_ltd,
                                              btt_id,
                                              kykk_tu_ngay,
                                              kykk_den_ngay,
                                              ctg_id,
                                              so_kkhai,
                                              so_dchinh,
                                              so_clech_dtnt,
                                              so_clech_cqt,
                                              ly_do_dchinh,
                                              ke_khai_sai)
                VALUES(qlt_xltk_dtl_seq.NEXTVAL,
                       p_TKhai_Exits_Id,
                       0,
                       v_Array_Pluc_01C(i).btt_id,
                       v_Array_Pluc_01C(i).kykk_tu_ngay,
                       v_Array_Pluc_01C(i).kykk_den_ngay,
                       v_Array_Pluc_01C(i).ctg_id,
                       v_Array_Pluc_01C(i).so_kkhai,
                       v_Array_Pluc_01C(i).so_dchinh,
                       v_Array_Pluc_01C(i).so_clech_dtnt,
                       v_Array_Pluc_01C(i).so_clech_cqt,
                       v_Array_Pluc_01C(i).ly_do_dchinh,
                       v_Array_Pluc_01C(i).ke_khai_sai);
            END LOOP;
            /*Thong tin bang phat sinh*/
            FOR i iN 1..v_Index LOOP
                --Danh sach tieu muc co tren to khai trong mode sua
                v_tmuc_exist := v_tmuc_exist ||','||v_Array_Psinh_Dtl(i).ma_tmuc;
                --Lay so phat sinh tren to khai truoc
                OPEN c_Psinh_TNguyen_Exist(v_Array_Psinh_Dtl(i).ma_muc,
                                           v_Array_Psinh_Dtl(i).ma_tmuc);
                FETCH c_Psinh_TNguyen_Exist INTO vc_Psinh_TNguyen_Exist;
                IF c_Psinh_TNguyen_Exist%FOUND THEN
                --Neu muc phat sinh da ton tai
                --IF NVL(vc_Psinh_TNguyen_Exist.tmt_ma_muc,'0') <> '0' THEN
                    IF vc_Psinh_TNguyen_Exist.thue_psinh <> v_Array_Psinh_Dtl(i).thue_psinh THEN
                        UPDATE qlt_psinh_tkhai
                        SET thue_psinh = v_Array_Psinh_Dtl(i).thue_psinh
                        WHERE (tkh_id = p_TKhai_Exits_Id)
                          AND (tkh_ltd = 0)
                          AND (tmt_ma_muc = vc_Psinh_TNguyen_Exist.tmt_ma_muc)
                          AND (tmt_ma_tmuc = vc_Psinh_TNguyen_Exist.tmt_ma_tmuc);
                    END IF;
                /* Neu muc chua ton tai */
                ELSE
                        INSERT INTO qlt_psinh_tkhai (id,
                                                     tkh_id,
                                                     tkh_ltd,
                                                     ccg_ma_cap,
                                                     ccg_ma_chuong,
                                                     lkn_ma_loai,
                                                     lkn_ma_khoan,
                                                     tmt_ma_muc,
                                                     tmt_ma_tmuc,
                                                     tmt_ma_thue,
                                                     thue_psinh)
                        VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                                p_TKhai_Exits_Id,
                                0,
                                p_Record_Dtnt.ma_cap,
                                p_Record_Dtnt.ma_chuong,
                                p_Record_Dtnt.ma_loai,
                                p_Record_Dtnt.ma_khoan,
                                v_Array_Psinh_Dtl(i).ma_muc,
                                v_Array_Psinh_Dtl(i).ma_tmuc,
                                v_Ma_Thue,
                                v_Array_Psinh_Dtl(i).thue_psinh);
                END IF;
                CLOSE c_Psinh_TNguyen_Exist;
            END LOOP;

            --Xoa nhung muc khong ton tai
            DELETE FROM qlt_psinh_tkhai
            WHERE (tkh_id = p_TKhai_Exits_Id)
            AND (tkh_ltd = 0)
            AND (INSTR(v_tmuc_exist,tmt_ma_tmuc) = 0);
        ELSE --To khai chua ton tai
            --Neu co an dinh
            IF (Fnc_AnDinh_Exits(p_Record_Of_Header.tin,
                                 p_Record_Of_Header.loai_tkhai,
                                 p_Record_Of_Header.kykk_tu_ngay,
                                 p_Record_Of_Header.kykk_den_ngay)) THEN
                v_Tthai_Tkhai := '3'; --To khai nop cham
            ELSE --Neu chua co an dinh
                v_HanNop := Fnc_Han_Nop(p_Record_Of_Header.loai_tkhai,
                                        'TK',
                                        p_Record_Of_Header.kykk_den_ngay);
                IF (p_Record_Of_Header.ngay_nop <= v_HanNop) THEN
                    --To khai dung han
                    v_Tthai_Tkhai := '1'; --To khai chinh thuc
                ELSE
                    --To khai khong dung han
                    v_Tthai_Tkhai := '3'; --To khai nop cham
                END IF;
            END IF;

            /* Thuc hien xu ly to khai */
            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

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
                                      co_loi,
                                      co_loi_ddanh,
                                      ghi_chu_loi,
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   v_Sai_SoHoc,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   NULL,  --p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat);

            /* Xu ly voi du lieu to khai Detail */
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_ttdb (id,
                                            tkh_id,
                                            tkh_ltd,
                                            btt_id,
                                            dvt_don_vi_tinh,
                                            so_luong,
                                            tong_tri_gia_ban,
                                            tong_tri_gia_tt_dtnt,
                                            thue_duoc_ktru,
                                            tsuat_dtnt,
                                            thue_phai_nop_dtnt,
                                            tong_tri_gia_tt_cqt,
                                            thue_phai_nop_cqt,
                                            ke_khai_sai,
                                            dchinh_tang_giam,
                                            so_tt,
                                            thue_pnop_tky_dtnt,
                                            thue_pnop_tky_cqt)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btt_id,
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).so_luong,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_ban,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_duoc_ktru,
                        v_Array_Tkhai_Dtl(i).tsuat_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_dtnt,
                        v_Array_Tkhai_Dtl(i).tong_tri_gia_tt_cqt,
                        v_Array_Tkhai_Dtl(i).thue_phai_nop_cqt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai,
                        v_Array_Tkhai_Dtl(i).dchinh_tang_giam,
                        v_Array_Tkhai_Dtl(i).so_tt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_tky_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_tky_cqt);
            END LOOP;
            /*Insert chitieu khong phat sinh tri gia tinh thue TTDB trong ky*/
            OPEN c_TKhai_TTDB(154);
            FETCH c_TKhai_TTDB INTO vc_TKhai_TTDB;
            CLOSE c_TKhai_TTDB;
            INSERT INTO qlt_tkhai_ttdb (id,
                            tkh_id,
                            tkh_ltd,
                            btt_id,
                            dvt_don_vi_tinh,
                            so_luong,
                            tong_tri_gia_ban,
                            tong_tri_gia_tt_dtnt,
                            thue_duoc_ktru,
                            tsuat_dtnt,
                            thue_phai_nop_dtnt,
                            tong_tri_gia_tt_cqt,
                            thue_phai_nop_cqt,
                            dchinh_tang_giam,
                            so_tt,
                            thue_pnop_tky_dtnt,
                            thue_pnop_tky_cqt)
            VALUES(qlt_xltk_dtl_seq.NEXTVAL,
                   v_Hdr_Id,
                   0,
                   154,
                   NULL,
                   DECODE(UPPER(vc_TKhai_TTDB.dvt_don_vi_tinh),'X',1,0),
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   0,
                   1,
                   0,
                   0);
                /*Insert phu luc to khai*/
                FOR i IN 1..v_Count_Pluc LOOP
                    INSERT INTO qlt_pluc_ttdb_01c(id,
                                                  tkh_id,
                                                  tkh_ltd,
                                                  btt_id,
                                                  kykk_tu_ngay,
                                                  kykk_den_ngay,
                                                  ctg_id,
                                                  so_kkhai,
                                                  so_dchinh,
                                                  so_clech_dtnt,
                                                  so_clech_cqt,
                                                  ly_do_dchinh,
                                                  ke_khai_sai)
                    VALUES(qlt_xltk_dtl_seq.NEXTVAL,
                           v_Hdr_Id,
                           0,
                           v_Array_Pluc_01C(i).btt_id,
                           v_Array_Pluc_01C(i).kykk_tu_ngay,
                           v_Array_Pluc_01C(i).kykk_den_ngay,
                           v_Array_Pluc_01C(i).ctg_id,
                           v_Array_Pluc_01C(i).so_kkhai,
                           v_Array_Pluc_01C(i).so_dchinh,
                           v_Array_Pluc_01C(i).so_clech_dtnt,
                           v_Array_Pluc_01C(i).so_clech_cqt,
                           v_Array_Pluc_01C(i).ly_do_dchinh,
                           v_Array_Pluc_01C(i).ke_khai_sai);
                END LOOP;
                /*Insert du lieu vao bang phat sinh */
                FOR i iN 1..v_Index LOOP
                        INSERT INTO qlt_psinh_tkhai (id,
                                                     tkh_id,
                                                     tkh_ltd,
                                                     ccg_ma_cap,
                                                     ccg_ma_chuong,
                                                     lkn_ma_loai,
                                                     lkn_ma_khoan,
                                                     tmt_ma_muc,
                                                     tmt_ma_tmuc,
                                                     tmt_ma_thue,
                                                     thue_psinh)
                        VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                                v_Hdr_Id,
                                0,
                                p_Record_Dtnt.ma_cap,
                                p_Record_Dtnt.ma_chuong,
                                p_Record_Dtnt.ma_loai,
                                p_Record_Dtnt.ma_khoan,
                                v_Array_Psinh_Dtl(i).ma_muc,
                                v_Array_Psinh_Dtl(i).ma_tmuc,
                                v_Ma_Thue,
                                v_Array_Psinh_Dtl(i).thue_psinh);
                 END LOOP;
            END IF;
        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);

        /*Ghi so nhan ho so*/
        IF v_Exist_Pluc THEN
            v_Ds_PLuc := '216,218';
        ELSE
            v_Ds_PLuc := '216';
        END IF;
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '25',
                         '02', -- To khai thue
                         v_Ds_PLuc);
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_TTDB'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Khainhg
Ngay lap: 19/04/2006
Noi dung: Thuc hien do du lieu to khai quyet toan thue tai nguyen vao trong CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai

********************************************************************************/
    PROCEDURE Prc_TKhai_Qtoan_TNguyen(p_Record_Of_Header Record_Hdr,
                                      p_Record_Dtnt Record_Dtnt,
                                      p_TKhai_Exits_Id NUMBER,
                                      p_Tthai_Tkhai VARCHAR2) IS

	    TYPE Record_Dtl_TNguyen IS RECORD(btn_id NUMBER(10,0),
                                          dvt_don_vi_tinh VARCHAR2(10),
                                          san_luong NUMBER(20,3),
                                          gia_ad_don_vi NUMBER(20,2),
                                          gia_tt_don_vi NUMBER(20,2),
                                          thue_suat NUMBER(5,2),
                                          thue_pnop_nam_dtnt NUMBER(20,2),
                                          thue_pnop_nam_cqt NUMBER(20,2),
                                          thue_psinh_nam_dtnt NUMBER(20,2),
                                          thue_psinh_nam_cqt NUMBER(20,2),
                                          thue_mgiam_nam_dtnt NUMBER(20,2),
                                          thue_mgiam_nam_cqt NUMBER(20,2),
                                          ke_khai_sai VARCHAR2(1));

        TYPE Record_Dtl_PSINH IS RECORD(ma_muc VARCHAR2(3),
                                        ma_tmuc VARCHAR2(2),
                                        so_kkhai_qtoan NUMBER(20,2),
                                        so_kkhai_tkhai NUMBER(20,2),
                                        so_clech NUMBER(20,2));

        TYPE Array_Tkhai_Dtl IS TABLE OF Record_Dtl_TNguyen INDEX BY BINARY_INTEGER;

        TYPE Array_Psinh_Dtl IS TABLE OF Record_Dtl_PSINH INDEX BY BINARY_INTEGER;

        v_Array_Tkhai_Dtl Array_Tkhai_Dtl;

        v_Array_Tkhai_Dtl_Tong Array_Tkhai_Dtl;

        v_Array_Psinh_Dtl Array_Psinh_Dtl;

        /* Lay chi tiet ke khai thue TNGUYEN */
        CURSOR c_TKhai_TNguyen IS
            SELECT *
            FROM rcv_v_tkhai_qtoan_tnguyen
            WHERE (hdr_id = p_Record_Of_Header.id)
              AND (btn_id IS NOT NULL);

        /* Lay dia diem khai thac, chi tieu tong */
        CURSOR c_Ctieu_Tong IS
            SELECT ddiem_kthac
                  ,thue_psinh_nam_dtnt
                  ,thue_mgiam_nam_dtnt
                  ,thue_pnop_nam_dtnt
            FROM rcv_v_tkhai_qtoan_tnguyen
            WHERE (hdr_id = p_Record_Of_Header.id);

        /* Tong hop so phat sinh */
        CURSOR c_Psinh_TNguyen IS
            SELECT  bthue.ma_muc,
                    bthue.ma_tmuc,
                    SUM(thue_psinh_nam_dtnt) so_kkhai_qtoan
            FROM rcv_v_tkhai_qtoan_tnguyen tkhai
               , rcv_map_ctieu_bthue bthue
               , qlt_dm_bthue_tnguyen dm
            WHERE (tkhai.btn_id IS NOT NULL)
              AND (tkhai.hdr_id = p_Record_Of_Header.id)
              AND (tkhai.btn_id = dm.id)
              AND (dm.ma = bthue.ma_ctieu)
              AND (bthue.loai_tkhai = '06')
              AND (dm.ngay_hl = (SELECT MAX(ngay_hl)
                                 FROM qlt_dm_bthue_tnguyen
                                 WHERE (ngay_hl< p_Record_Of_Header.kylb_tu_ngay)))
            GROUP BY bthue.ma_muc,bthue.ma_tmuc;
        /* Lay so phat sinh tren to khai qtoan TNGUYEN da ton tai */
        CURSOR c_Psinh_Qtoan_Exist(p_ma_muc VARCHAR2,
                                   p_ma_tmuc VARCHAR2) IS
           SELECT *
           FROM qlt_psinh_qtoan
           WHERE (quh_id = p_TKhai_Exits_Id)
             AND (quh_ltd = 0)
             AND (tmt_ma_muc = p_ma_muc)
             AND (tmt_ma_tmuc = p_ma_tmuc);

        /* Lay ma loai thue */
        CURSOR c_Loai_Thue IS
            SELECT lte_ma_thue
            FROM qlt_dm_qtoan
            WHERE (ma = p_Record_Of_Header.loai_tkhai);

        vc_Loai_Thue c_Loai_Thue%ROWTYPE;
        vc_Ctieu_Tong c_Ctieu_Tong%ROWTYPE;
        vc_Psinh_Qtoan_Exist c_Psinh_Qtoan_Exist%ROWTYPE;
        v_HanNop DATE;
        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_Ma_Thue VARCHAR(2);
        v_Thue_PSinh NUMBER(20,2);
        v_Check BOOLEAN := TRUE;
        v_Count NUMBER(10,0) := 0;
        v_Index NUMBER(10,0) :=0;
        v_Sai_SoHoc VARCHAR2(1) := NULL;
        v_Thue_Psinh_Nam_Cqt NUMBER := 0;
        v_Thue_PNop_Nam_Cqt NUMBER := 0;
        v_Exist_tmuc VARCHAR2(100);

    BEGIN
        /* Lay ma thue */
        OPEN c_Loai_Thue;
        FETCH c_Loai_Thue INTO vc_Loai_Thue;
        IF (c_Loai_Thue%FOUND) THEN
            v_Ma_Thue := vc_Loai_Thue.lte_ma_thue;
        END IF;
        CLOSE c_Loai_Thue;
        --lay dia diem khai thac, ctieu tong
        OPEN c_Ctieu_Tong;
        FETCH c_Ctieu_Tong INTO vc_Ctieu_Tong;
        CLOSE c_Ctieu_Tong;

        /* Luu cac gia tri cua to khai TNGUYEN truc tiep ra mang */
        FOR vc_TKhai_TNguyen IN c_TKhai_TNguyen LOOP
            v_Count := v_Count + 1;
            v_Array_Tkhai_Dtl(v_Count).btn_id := vc_TKhai_TNguyen.btn_id;
            v_Array_Tkhai_Dtl(v_Count).dvt_don_vi_tinh := vc_TKhai_TNguyen.don_vi_tinh;
            v_Array_Tkhai_Dtl(v_Count).san_luong := vc_TKhai_TNguyen.san_luong;
            v_Array_Tkhai_Dtl(v_Count).gia_ad_don_vi := vc_TKhai_TNguyen.gia_tt_don_vi;
            v_Array_Tkhai_Dtl(v_Count).gia_tt_don_vi := vc_TKhai_TNguyen.gia_ad_don_vi;
            v_Array_Tkhai_Dtl(v_Count).thue_suat := vc_TKhai_TNguyen.tsuat_dtnt;
            v_Array_Tkhai_Dtl(v_Count).thue_pnop_nam_dtnt := NVL(vc_TKhai_TNguyen.thue_pnop_nam_dtnt,0);
            v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_dtnt := NVL(vc_TKhai_TNguyen.thue_psinh_nam_dtnt,0);
            v_Array_Tkhai_Dtl(v_Count).thue_mgiam_nam_dtnt := NVL(vc_TKhai_TNguyen.thue_mgiam_nam_dtnt,0);

            /*
                - Tinh so co quan thue.
                - Neu Gia tinh thue don vi tai nguyen = 0
                    -> Thue psinh trong ky = San luong
                                           * muc thue an dinh tren 1 dv tai nguyen
                  Else Thue psinh trong ky = San luong
                                           * Gia tinh thue don vi
            */
            IF NVL(v_Array_Tkhai_Dtl(v_Count).gia_ad_don_vi,0) =0 THEN
                v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_cqt := v_Array_Tkhai_Dtl(v_Count).san_luong
                                                               * v_Array_Tkhai_Dtl(v_Count).gia_tt_don_vi;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_cqt := Round(v_Array_Tkhai_Dtl(v_Count).san_luong
                                                               * v_Array_Tkhai_Dtl(v_Count).gia_ad_don_vi
                                                               * v_Array_Tkhai_Dtl(v_Count).thue_suat/100);
            END IF;
            v_Array_Tkhai_Dtl(v_Count).thue_pnop_nam_cqt := v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_cqt
                                                          - v_Array_Tkhai_Dtl(v_Count).thue_mgiam_nam_dtnt;
            /*
                - Kiem tra so dtnt va so cqt
                - Neu thue psinh tky Dtnt <> thue psinh tky cqt
                    or thue pnop dtnt <> thue pnop cqt (kiem tra theo nguong) -> ke khai sai = 'Y'
            */
            IF (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_dtnt,
                               v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_cqt,
                               p_Record_Of_Header.kylb_tu_ngay))
                AND (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl(v_Count).thue_pnop_nam_dtnt,
                                    v_Array_Tkhai_Dtl(v_Count).thue_pnop_nam_cqt,
                                    p_Record_Of_Header.kylb_tu_ngay)) THEN
                v_Array_Tkhai_Dtl(v_Count).ke_khai_sai := NULL;
            ELSE
                v_Array_Tkhai_Dtl(v_Count).ke_khai_sai := 'Y';
                v_Sai_SoHoc := 'Y';
            END IF;
            v_Thue_Psinh_Nam_Cqt := v_Thue_Psinh_Nam_Cqt
                                  + NVL(v_Array_Tkhai_Dtl(v_Count).thue_psinh_nam_cqt,0);
            v_Thue_PNop_Nam_Cqt := v_Thue_PNop_Nam_Cqt
                                  + NVL(v_Array_Tkhai_Dtl(v_Count).thue_pnop_nam_cqt,0);
        END LOOP;
        /* Tinh chi tieu tong DTNT */
        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_cqt := v_Thue_Psinh_Nam_Cqt;
        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_cqt := v_Thue_PNop_Nam_Cqt;
        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_dtnt := NVL(vc_Ctieu_Tong.thue_psinh_nam_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).thue_mgiam_nam_dtnt := NVL(vc_Ctieu_Tong.thue_mgiam_nam_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_dtnt := NVL(vc_Ctieu_Tong.thue_pnop_nam_dtnt,0);
        v_Array_Tkhai_Dtl_Tong(1).btn_id := 68;
        /*
            - Kiem tra so dtnt va so cqt
            - Neu thue psinh tky Dtnt <> thue psinh tky cqt
                or thue pnop dtnt <> thue pnop cqt (kiem tra theo nguong) -> ke khai sai = 'Y'
        */
        IF (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_dtnt,
                           v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_cqt,
                           p_Record_Of_Header.kylb_tu_ngay))
            AND (Fnc_Ktra_Ctieu(v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_dtnt,
                                v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_cqt,
                                p_Record_Of_Header.kylb_tu_ngay)) THEN
            v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai := NULL;
        ELSE
            v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai := 'Y';
            v_Sai_SoHoc := 'Y';
        END IF;
        /*
            - Tinh so phat sinh
            - So kkhai qtoan: so ke khai tren to khai quyet toan
            - So kkhai tkhai: so ke khai tren to khai 12 thang
            - So chenh lech: so kkhai qtoan - so kkhai tkhai
        */
        FOR vc_Psinh_TNguyen IN c_Psinh_TNguyen LOOP
            v_Index := v_Index +1;
            v_Array_Psinh_Dtl(v_Index).ma_muc := vc_Psinh_TNguyen.ma_muc;
            v_Array_Psinh_Dtl(v_Index).ma_tmuc := vc_Psinh_TNguyen.ma_tmuc;
            v_Array_Psinh_Dtl(v_Index).so_kkhai_qtoan := vc_Psinh_TNguyen.so_kkhai_qtoan;
            v_Array_Psinh_Dtl(v_Index).so_kkhai_tkhai := Fnc_So_KKhai_Tnguyen(p_Record_Of_Header.tin,
                                                                              vc_Psinh_TNguyen.ma_muc,
                                                                              vc_Psinh_TNguyen.ma_tmuc,
                                                                              v_Ma_Thue,
                                                                              p_Record_Of_Header.kykk_tu_ngay);

            v_Array_Psinh_Dtl(v_Index).so_clech := NVL(v_Array_Psinh_Dtl(v_Index).so_kkhai_qtoan,0)
                                                 - NVL(v_Array_Psinh_Dtl(v_Index).so_kkhai_tkhai,0);
        END LOOP;

        /*Neu da ton tai to khai trong ky ke khai*/
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            Qlt_Pck_Gdich.Prc_Lay_Thamso_Qtoan('08');
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;
            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_QTOAN_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);
            --Xoa du lieu trong tat ca cac bang voi dieu kien ltd =0
            DELETE FROM qlt_qtoan_tnguyen
            WHERE (quh_id = p_TKhai_Exits_Id)
              AND (quh_ltd = 0);

            /*Thong tin Header*/
            UPDATE qlt_qtoan_hdr
            SET ngay_nop = p_Record_Of_Header.ngay_nop,
                kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                tthai = '4', --To khai thay the
                loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat,
                --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                so_tt_tk = p_Record_Of_Header.so_tt_tk,
                ddiem_kthac = vc_Ctieu_Tong.ddiem_kthac
            WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);

            /*Thong tin Detail*/
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_qtoan_tnguyen (id,
                                               quh_id,
                                               quh_ltd,
                                               btn_id,
                                               so_tt,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_ad_don_vi,
                                               gia_tt_don_vi,
                                               thue_suat,
                                               thue_pnop_nam_dtnt,
                                               thue_pnop_nam_cqt,
                                               thue_psinh_nam_dtnt,
                                               thue_psinh_nam_cqt,
                                               thue_mgiam_nam_dtnt,
                                               thue_mgiam_nam_cqt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        p_TKhai_Exits_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btn_id,
                        i,--so_tt
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).san_luong,
                        v_Array_Tkhai_Dtl(i).gia_ad_don_vi,
                        v_Array_Tkhai_Dtl(i).gia_tt_don_vi,
                        v_Array_Tkhai_Dtl(i).thue_suat,
                        v_Array_Tkhai_Dtl(i).thue_pnop_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_nam_cqt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_nam_cqt,
                        v_Array_Tkhai_Dtl(i).thue_mgiam_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_mgiam_nam_cqt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai);
            END LOOP;
            /* Insert chi tieu tong */
                INSERT INTO qlt_qtoan_tnguyen (id,
                                               quh_id,
                                               quh_ltd,
                                               btn_id,
                                               so_tt,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_ad_don_vi,
                                               gia_tt_don_vi,
                                               thue_suat,
                                               thue_pnop_nam_dtnt,
                                               thue_pnop_nam_cqt,
                                               thue_psinh_nam_dtnt,
                                               thue_psinh_nam_cqt,
                                               thue_mgiam_nam_dtnt,
                                               thue_mgiam_nam_cqt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        p_TKhai_Exits_Id,
                        0,
                        v_Array_Tkhai_Dtl_Tong(1).btn_id,
                        NULL,--so_tt
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_mgiam_nam_dtnt,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai);
            /*Thong tin bang phat sinh*/
            FOR i iN 1..v_Index LOOP
                --Danh sach muc phat sinh tren to khai qtoan.
                v_Exist_tmuc := v_Exist_tmuc||','||v_Array_Psinh_Dtl(i).ma_tmuc;
                 --Lay so phat sinh tren to khai truoc
                OPEN c_Psinh_Qtoan_Exist(v_Array_Psinh_Dtl(i).ma_muc,
                                         v_Array_Psinh_Dtl(i).ma_tmuc);
                FETCH c_Psinh_Qtoan_Exist INTO vc_Psinh_Qtoan_Exist;
                IF c_Psinh_Qtoan_Exist%FOUND THEN
                --Neu da ton tai muc phat sinh
                --IF NVL(vc_Psinh_Qtoan_Exist.tmt_ma_muc,'0') <> '0' THEN
                    IF vc_Psinh_Qtoan_Exist.so_kkhai_qtoan <> v_Array_Psinh_Dtl(i).so_kkhai_qtoan THEN
                        UPDATE qlt_psinh_qtoan
                        SET so_kkhai_qtoan = v_Array_Psinh_Dtl(i).so_kkhai_qtoan
                        WHERE (quh_id = p_TKhai_Exits_Id)
                          AND (quh_ltd =0)
                          AND (tmt_ma_muc = vc_Psinh_Qtoan_Exist.tmt_ma_muc)
                          AND (tmt_ma_tmuc = vc_Psinh_Qtoan_Exist.tmt_ma_tmuc);
                    ELSIF vc_Psinh_Qtoan_Exist.so_kkhai_tkhai <> v_Array_Psinh_Dtl(i).so_kkhai_tkhai THEN
                        UPDATE qlt_psinh_qtoan
                        SET so_kkhai_tkhai = v_Array_Psinh_Dtl(i).so_kkhai_tkhai
                        WHERE (quh_id = p_TKhai_Exits_Id)
                          AND (quh_ltd = 0)
                          AND (tmt_ma_muc = vc_Psinh_Qtoan_Exist.tmt_ma_muc)
                          AND (tmt_ma_tmuc = vc_Psinh_Qtoan_Exist.tmt_ma_tmuc);
                    ELSIF vc_Psinh_Qtoan_Exist.so_clech <> v_Array_Psinh_Dtl(i).so_clech THEN
                        UPDATE qlt_psinh_qtoan
                        SET so_clech = v_Array_Psinh_Dtl(i).so_clech
                        WHERE (quh_id = p_TKhai_Exits_Id)
                          AND (quh_ltd = 0)
                          AND (tmt_ma_muc = vc_Psinh_Qtoan_Exist.tmt_ma_muc)
                          AND (tmt_ma_tmuc = vc_Psinh_Qtoan_Exist.tmt_ma_tmuc);
                    END IF;
                --Neu muc chua ton tai
                ELSE
                    INSERT INTO qlt_psinh_qtoan (id,
                                                 quh_id,
                                                 quh_ltd,
                                                 ccg_ma_cap,
                                                 ccg_ma_chuong,
                                                 lkn_ma_loai,
                                                 lkn_ma_khoan,
                                                 tmt_ma_muc,
                                                 tmt_ma_tmuc,
                                                 tmt_ma_thue,
                                                 so_kkhai_qtoan,
                                                 so_kkhai_tkhai,
                                                 so_clech)
                    VALUES (qlt_xltk_hdr_seq.NEXTVAL,
                            p_TKhai_Exits_Id,
                            0,
                            p_Record_Dtnt.ma_cap,
                            p_Record_Dtnt.ma_chuong,
                            p_Record_Dtnt.ma_loai,
                            p_Record_Dtnt.ma_khoan,
                            v_Array_Psinh_Dtl(i).ma_muc,
                            v_Array_Psinh_Dtl(i).ma_tmuc,
                            v_Ma_Thue,
                            v_Array_Psinh_Dtl(i).so_kkhai_qtoan,
                            v_Array_Psinh_Dtl(i).so_kkhai_tkhai,
                            v_Array_Psinh_Dtl(i).so_clech);
                END IF;
                CLOSE c_Psinh_Qtoan_Exist;
            END LOOP;
            --Xoa nhung muc phat sinh khong ton tai
            DELETE FROM qlt_psinh_qtoan
            WHERE (quh_id = p_TKhai_Exits_Id)
              AND (quh_ltd = 0)
              AND (INSTR(v_Exist_tmuc,tmt_ma_tmuc) = 0);

        ELSE --To khai chua ton tai
            v_Tthai_Tkhai := '1'; --trang thai chinh thuc
            --Thuc hien xu ly to khai
            Qlt_Pck_Gdich.Prc_Lay_Thamso_Qtoan('08');
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Xu ly voi du lieu to khai Header
            INSERT INTO qlt_qtoan_hdr(id,
                                      ltd,
                                      tin,
                                      ten_dtnt,
                                      cqt_ma_cqt,
                                      hun_ma_tinh,
                                      hun_ma_huyen,
                                      dia_chi,
                                      ma_phong,
                                      ma_can_bo,
                                      dqt_ma,
                                      ngay_nop,
                                      kylb_tu_ngay,
                                      kylb_den_ngay,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      tthai,
                                      loi_so_hoc,
                                      loi_ddanh,
                                      ghi_chu_loi,
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat,
                                      ddiem_kthac)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   v_Sai_SoHoc,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   NULL,  --p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat,
                   vc_Ctieu_Tong.ddiem_kthac);

            --Xu ly voi du lieu to khai Detail
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_qtoan_tnguyen (id,
                                               quh_id,
                                               quh_ltd,
                                               btn_id,
                                               so_tt,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_ad_don_vi,
                                               gia_tt_don_vi,
                                               thue_suat,
                                               thue_pnop_nam_dtnt,
                                               thue_pnop_nam_cqt,
                                               thue_psinh_nam_dtnt,
                                               thue_psinh_nam_cqt,
                                               thue_mgiam_nam_dtnt,
                                               thue_mgiam_nam_cqt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Tkhai_Dtl(i).btn_id,
                        i, --so_tt
                        v_Array_Tkhai_Dtl(i).dvt_don_vi_tinh,
                        v_Array_Tkhai_Dtl(i).san_luong,
                        v_Array_Tkhai_Dtl(i).gia_ad_don_vi,
                        v_Array_Tkhai_Dtl(i).gia_tt_don_vi,
                        v_Array_Tkhai_Dtl(i).thue_suat,
                        v_Array_Tkhai_Dtl(i).thue_pnop_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_pnop_nam_cqt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_psinh_nam_cqt,
                        v_Array_Tkhai_Dtl(i).thue_mgiam_nam_dtnt,
                        v_Array_Tkhai_Dtl(i).thue_mgiam_nam_cqt,
                        v_Array_Tkhai_Dtl(i).ke_khai_sai);
            END LOOP;
            /* Insert chi tieu tong */
                INSERT INTO qlt_qtoan_tnguyen (id,
                                               quh_id,
                                               quh_ltd,
                                               btn_id,
                                               so_tt,
                                               dvt_don_vi_tinh,
                                               san_luong,
                                               gia_ad_don_vi,
                                               gia_tt_don_vi,
                                               thue_suat,
                                               thue_pnop_nam_dtnt,
                                               thue_pnop_nam_cqt,
                                               thue_psinh_nam_dtnt,
                                               thue_psinh_nam_cqt,
                                               thue_mgiam_nam_dtnt,
                                               thue_mgiam_nam_cqt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Tkhai_Dtl_Tong(1).btn_id,
                        NULL,--so_tt
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_pnop_nam_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_dtnt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_psinh_nam_cqt,
                        v_Array_Tkhai_Dtl_Tong(1).thue_mgiam_nam_dtnt,
                        NULL,
                        v_Array_Tkhai_Dtl_Tong(1).ke_khai_sai);

            --Insert du lieu vao bang phat sinh
            FOR i IN 1..v_Index LOOP
                INSERT INTO qlt_psinh_qtoan (id,
                                             quh_id,
                                             quh_ltd,
                                             ccg_ma_cap,
                                             ccg_ma_chuong,
                                             lkn_ma_loai,
                                             lkn_ma_khoan,
                                             tmt_ma_muc,
                                             tmt_ma_tmuc,
                                             tmt_ma_thue,
                                             so_kkhai_qtoan,
                                             so_kkhai_tkhai,
                                             so_clech)
                VALUES (qlt_xltk_hdr_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        p_Record_Dtnt.ma_cap,
                        p_Record_Dtnt.ma_chuong,
                        p_Record_Dtnt.ma_loai,
                        p_Record_Dtnt.ma_khoan,
                        v_Array_Psinh_Dtl(i).ma_muc,
                        v_Array_Psinh_Dtl(i).ma_tmuc,
                        v_Ma_Thue,
                        v_Array_Psinh_Dtl(i).so_kkhai_qtoan,
                        v_Array_Psinh_Dtl(i).so_kkhai_tkhai,
                        v_Array_Psinh_Dtl(i).so_clech);
            END LOOP;
        END IF;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        /*Ghi so nhan ho so*/
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '08', --Quyet toan thue TNGUYEN
                         '03', --To khai quyet toan
                         '226'); --Ma phu luc
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_Qtoan_Tnguyen'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 19/01/2006
Muc dich: Thuc hien do du lieu to khai quyet toan TNDN nam vao cac bang
          trong CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai
*******************************************************************************/
    PROCEDURE Prc_TKhai_QToan_TNDN_Nam(p_Record_Of_Header Record_Hdr,
                                       p_Record_Dtnt Record_Dtnt,
                                       p_TKhai_Exits_Id NUMBER,
                                       p_Tthai_Tkhai VARCHAR2) IS
        --Lay so lieu to khai quyet toan
        CURSOR c_TKhai_QToan_TNDN_Nam IS
            SELECT *
            FROM rcv_v_tkhai_qtoan_tndn qtoan
            WHERE (qtoan.hdr_id = p_Record_Of_Header.id)
            ORDER BY qtoan.ky_hieu_ctieu, qtoan.so_tt;
        vc_TKhai_QToan_TNDN_Nam c_TKhai_QToan_TNDN_Nam%ROWTYPE;

        --Lay ma thue
        CURSOR c_Ma_Thue IS
        	SELECT lte_ma_thue
        	FROM qlt_dm_qtoan
        	WHERE (ma = p_Record_Of_Header.loai_tkhai);

        --Lay cac nam chuyen lo
        CURSOR c_Nam_ChuyenLo(p_Loai_PLuc VARCHAR2) IS
            SELECT pluc_01AB.lo_chuyen_01 nam_chuyen_1,
                   pluc_01AB.lo_chuyen_02 nam_chuyen_2,
                   pluc_01AB.lo_chuyen_03 nam_chuyen_3,
                   pluc_01AB.lo_chuyen_04 nam_chuyen_4,
                   pluc_01AB.lo_chuyen_05 nam_chuyen_5,
                   pluc_01AB.lo_chuyen_06 nam_chuyen_6
            FROM rcv_v_pluc_qtoan_tndn_01AB pluc_01AB
            WHERE (pluc_01AB.hdr_id = p_Record_Of_Header.id)
              AND (pluc_01AB.loai_dlieu = p_Loai_PLuc)
              AND (pluc_01AB.nam_psinh IS NULL);
        vc_Nam_ChuyenLo c_Nam_ChuyenLo%ROWTYPE;

        --Lay du lieu phu luc 01A
        CURSOR c_PLuc_01A IS
            SELECT *
            FROM rcv_v_pluc_qtoan_tndn_01AB
            WHERE (hdr_id = p_Record_Of_Header.id)
              AND (loai_dlieu = '0302')
              AND (nam_psinh IS NOT NULL)
            ORDER BY loai_dlieu, so_tt;

        vc_PLuc_01A c_PLuc_01A%ROWTYPE;
        --Lay du lieu phu luc 01B
        CURSOR c_PLuc_01B IS
            SELECT *
            FROM rcv_v_pluc_qtoan_tndn_01AB
            WHERE (hdr_id = p_Record_Of_Header.id)
              AND (loai_dlieu = '0316')
              AND (nam_psinh IS NOT NULL)
            ORDER BY loai_dlieu, so_tt;

        vc_PLuc_01B c_PLuc_01B%ROWTYPE;
        --Lay du lieu tu phu luc 02 den phu luc 13
        CURSOR c_PLuc_02_13 IS
            SELECT *
            FROM rcv_v_pluc_qtoan_tndn_02_13
            WHERE (hdr_id = p_Record_Of_Header.id)
            ORDER BY loai_dlieu, row_id, ky_hieu;
        vc_PLuc_02_13 c_PLuc_02_13%ROWTYPE;

        --Lay du lieu phu luc 14
        CURSOR c_PLuc_14 IS
            SELECT *
            FROM rcv_v_pluc_qtoan_tndn_14
            WHERE (hdr_id = p_Record_Of_Header.id);

         vc_PLuc_14 c_PLuc_14%ROWTYPE;

        /* Khai bao phuc vu cho ghi so nhan ho so */
        CURSOR c_PLuc_Exist_02_13 IS
            SELECT DISTINCT pluc.loai
            FROM rcv_v_pluc_qtoan_tndn_02_13 rcv,
                 qlt_dm_ctieu_pluc_tndn pluc
            WHERE (hdr_id = p_Record_Of_Header.id)
            AND pluc.id = rcv.dcp_id;

        v_Ds_Pluc VARCHAR2(200):= NULL;

        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_PLuc01A_Id NUMBER(10);
        v_PLuc01B_Id NUMBER(10);
        v_Ma_Thue VARCHAR2(2);
        v_So_KKhai_QToan NUMBER(20,2) := 0;
        v_So_KKhai_TKhai NUMBER(20,2) := 0;
        v_So_Clech NUMBER(20,2) := 0;
        v_Count NUMBER(10) := 0;
        v_HanNop DATE;

    BEGIN
        --Xu ly so thue phat sinh
        OPEN c_Ma_Thue;
        FETCH c_Ma_Thue INTO v_Ma_Thue;
        CLOSE c_Ma_Thue;
        --Lay so ke khai cua doi tuong trong 12 thang
        v_So_KKhai_TKhai := Fnc_So_KKhai_TKhai_Quy (p_Record_Of_Header.tin,
            									    '002',
            									    '02',
            									    v_Ma_Thue,
            									    p_Record_Of_Header.kykk_tu_ngay);

        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            --Neu da ton tai to khai trong ky ke khai
            v_Hdr_Id := p_TKhai_Exits_Id;
            --Gan tham so sinh giao dich
        	Qlt_Pck_Gdich.Prc_Lay_Thamso_Qtoan(p_Record_Of_Header.loai_tkhai);
        	Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
        	--Thuc hien backup to khai
   		    Qlt_Pck_Tkhai.Prc_Backup_TKhai('QLT_QTOAN_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           v_Hdr_Id);
            --Cap nhat thong tin Header
            UPDATE qlt_qtoan_hdr
            SET ngay_nop = p_Record_Of_Header.ngay_nop,
                kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                tthai = '4', --To khai thay the
                loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat,
                --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                so_tt_tk = p_Record_Of_Header.so_tt_tk
            WHERE (id = v_Hdr_Id)
              AND (ltd = 0);

            --Bo cac tai lieu di kem to khai quyet toan
            UPDATE qlt_pluc_tlieu_qtoan_tndn
            SET co_khong = NULL
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            --Cap nhat thong tin Detail
            FOR vc_TKhai_QToan_TNDN_Nam IN c_TKhai_QToan_TNDN_Nam LOOP
                IF (vc_TKhai_QToan_TNDN_Nam.ky_hieu_ctieu = 'A') THEN
                    --Du lieu Detail quyet toan
                    UPDATE qlt_qtoan_dtl
                    SET so_dtnt = TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0)),
                        so_cqt = TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0)),
                        ke_khai_sai = NULL
                    WHERE (quh_id = v_Hdr_Id)
                      AND (quh_ltd = 0)
                      AND (ctq_id = vc_TKhai_QToan_TNDN_Nam.ctq_id);

                    IF (vc_TKhai_QToan_TNDN_Nam.ctq_id = 148) THEN
                        v_So_KKhai_QToan := TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0));
                    END IF;
                ELSE
                    --Cap nhat lai cac tai lieu di kem to khai quyet toan
                    UPDATE qlt_pluc_tlieu_qtoan_tndn
                    SET co_khong = DECODE(vc_TKhai_QToan_TNDN_Nam.so_dtnt,'x','Y',NULL)
                    WHERE (quh_id = v_Hdr_Id)
                      AND (quh_ltd = 0)
                      AND (dmtl_id = vc_TKhai_QToan_TNDN_Nam.ctq_id);
                END IF;
            END LOOP;

            --Tinh so chenh lech quyet toan
            v_So_Clech := v_So_KKhai_QToan - v_So_KKhai_TKhai;

            UPDATE qlt_psinh_qtoan
            SET so_kkhai_qtoan = v_So_KKhai_QToan,
                so_kkhai_tkhai = v_So_KKhai_TKhai,
                so_clech = v_So_Clech
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            --Xoa cac phu luc di kem to khai quyet toan
            --Xoa phu luc 01a, 01b
            DELETE FROM qlt_pluc_qtoan_tndn_01a
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            DELETE FROM qlt_pluc_qtoan_tndn_01b
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            --Xoa cac phu luc tu 2 den 13
            DELETE FROM qlt_pluc_qtoan_tndn_a
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            DELETE FROM qlt_pluc_qtoan_tndn_b
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);

            --Xoa phu luc 14
            DELETE FROM qlt_pluc_qtoan_tndn_14
            WHERE (quh_id = v_Hdr_Id)
              AND (quh_ltd = 0);
        ELSE
            v_Tthai_Tkhai := '1'; --To khai chinh thuc
            --Neu chua ton tai to khai trong ky ke khai
        	SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Gan tham so sinh giao dich
        	Qlt_Pck_Gdich.Prc_Lay_Thamso_Qtoan(p_Record_Of_Header.loai_tkhai);
        	Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);

        	--Thong tin Header quyet toan
            INSERT INTO qlt_qtoan_hdr(id,
                                      ltd,
                                      tin,
                                      ten_dtnt,
                                      cqt_ma_cqt,
                                      hun_ma_huyen,
                                      hun_ma_tinh,
                                      ma_phong,
                                      ma_can_bo,
                                      dia_chi,
                                      nganh_kdoanh,
                                      dien_thoai,
                                      fax,
                                      email,
                                      dqt_ma,
                                      ngay_nop,
                                      kylb_tu_ngay,
                                      kylb_den_ngay,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      tthai,
                                      loi_so_hoc,
                                      loi_ddanh,
                                      ghi_chu_loi,
                                      nguoi_cap_nhat,
                                      so_hieu_tep,
                                      so_tt_tk)
            VALUES (v_Hdr_Id,
                    0,
                    p_Record_Of_Header.tin,
                    p_Record_Of_Header.ten_dtnt,
                    p_Record_Dtnt.ma_cqt,
                    p_Record_Dtnt.ma_huyen,
                    p_Record_Dtnt.ma_tinh,
                    p_Record_Dtnt.ma_phong,
                    p_Record_Dtnt.ma_canbo,
                    p_Record_Of_Header.dia_chi,
                    NULL,
                    p_Record_Dtnt.dien_thoai,
                    p_Record_Dtnt.fax,
                    p_Record_Dtnt.email,
                    p_Record_Of_Header.loai_tkhai,
                    p_Record_Of_Header.ngay_nop,
                    p_Record_Of_Header.kylb_tu_ngay,
                    p_Record_Of_Header.kylb_den_ngay,
                    p_Record_Of_Header.kykk_tu_ngay,
                    p_Record_Of_Header.kykk_den_ngay,
                    v_Tthai_Tkhai, --Trang thai to khai
                    NULL,
                    p_Record_Of_Header.co_loi_ddanh,
                    p_Record_Of_Header.ghi_chu_loi,
                    p_Record_Of_Header.nguoi_cap_nhat,
                    NULL,  --p_Record_Of_Header.so_hieu_tep,
                    p_Record_Of_Header.so_tt_tk);

            FOR vc_TKhai_QToan_TNDN_Nam IN c_TKhai_QToan_TNDN_Nam LOOP
                IF (vc_TKhai_QToan_TNDN_Nam.ky_hieu_ctieu = 'A') THEN
                    --Du lieu Detail quyet toan
                    INSERT INTO qlt_qtoan_dtl (id,
                                               quh_id,
                                               quh_ltd,
                                               ctq_id,
                                               so_tt,
                                               so_dtnt,
                                               so_cqt,
                                               ke_khai_sai)
                    VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                            v_Hdr_Id,
                            0,
                            vc_TKhai_QToan_TNDN_Nam.ctq_id,
                            vc_TKhai_QToan_TNDN_Nam.so_tt,
                            TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0)),
                            TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0)),
                            NULL);

                    IF (vc_TKhai_QToan_TNDN_Nam.ctq_id = 148) THEN
                        v_so_kkhai_qtoan := TO_NUMBER(NVL(vc_TKhai_QToan_TNDN_Nam.so_dtnt,0));
                    END IF;
                ELSE
                    --Tai lieu di kem to khai quyet toan
                    INSERT INTO qlt_pluc_tlieu_qtoan_tndn(id,
                                                          quh_id,
                                                          quh_ltd,
                                                          dmtl_id,
                                                          so_tt,
                                                          co_khong)
                    VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                            v_Hdr_Id,
                            0,
                            vc_TKhai_QToan_TNDN_Nam.ctq_id,
                            vc_TKhai_QToan_TNDN_Nam.so_tt,
                            DECODE(vc_TKhai_QToan_TNDN_Nam.so_dtnt,'x','Y',NULL));
                END IF;
            END LOOP;

            --Tinh toan so chenh lech
            v_So_Clech := v_So_KKhai_QToan - v_So_KKhai_TKhai;
            --Dua du lieu vao bang phat sinh
            INSERT INTO qlt_psinh_qtoan (id,
                                         quh_id,
                                         quh_ltd,
                                         ccg_ma_cap,
                                         ccg_ma_chuong,
                                         lkn_ma_loai,
                                         lkn_ma_khoan,
                                         tmt_ma_muc,
                                         tmt_ma_tmuc,
                                         tmt_ma_thue,
                                         so_kkhai_qtoan,
                                         so_kkhai_tkhai,
                                         so_clech)
            VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    p_Record_Dtnt.ma_cap,
                    p_Record_Dtnt.ma_chuong,
                    p_Record_Dtnt.ma_loai,
                    p_Record_Dtnt.ma_khoan,
                    '002',
                    '02',
                    v_Ma_Thue,
                    v_So_KKhai_QToan,
                    v_So_KKhai_TKhai,
                    v_So_Clech);
        END IF;

        ---Xu ly cac phu luc di kem to khai quyet toan---

        --Xu ly du lieu phu luc 01A
        OPEN c_Nam_ChuyenLo('0302');
        FETCH c_Nam_ChuyenLo INTO vc_Nam_ChuyenLo;
        CLOSE c_Nam_ChuyenLo;

        FOR vc_PLuc_01A IN c_PLuc_01A LOOP
            IF (vc_PLuc_01A.so_tt < 8 AND vc_PLuc_01A.so_tt > 1) THEN
                IF ((vc_PLuc_01A.so_psinh <> 0) OR ((vc_PLuc_01A.so_psinh = 0) AND(vc_PLuc_01A.lo_chuyen_01 <> 0 OR
                                                                                   vc_PLuc_01A.lo_chuyen_02 <> 0 OR
                                                                                   vc_PLuc_01A.lo_chuyen_03 <> 0 OR
                                                                                   vc_PLuc_01A.lo_chuyen_04 <> 0 OR
                                                                                   vc_PLuc_01A.lo_chuyen_05 <> 0 OR
                                                                                   vc_PLuc_01A.lo_chuyen_06 <> 0))) THEN
                    SELECT qlt_xltk_dtl_seq.NEXTVAL INTO v_PLuc01A_Id FROM dual;
                    INSERT INTO qlt_pluc_qtoan_tndn_01a (id,
                                                         quh_id,
                                                         quh_ltd,
                                                         nam_psinh,
                                                         lo_psinh,
                                                         loai)
                    VALUES (v_PLuc01A_Id,
                            v_Hdr_Id,
                            0,
                            TO_DATE(vc_PLuc_01A.nam_psinh,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.so_psinh,0)),
                            '01');
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_01 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_1,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_01,0)));
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_02 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_2,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_02,0)));
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_03 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_3,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_03,0)));
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_04 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_4,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_04,0)));
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_05 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_5,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_05,0)));
                END IF;

                IF (vc_PLuc_01A.lo_chuyen_06 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01A_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_6,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_06,0)));
                END IF;
            ELSE
                INSERT INTO qlt_pluc_qtoan_tndn_01b (id,
                                                     quh_id,
                                                     quh_ltd,
                                                     nam_psinh,
                                                     lo_psinh,
                                                     lo_chuyen_truoc,
                                                     lo_chuyen_tky,
                                                     lo_chuyen_sau,
                                                     loai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        TO_DATE(vc_PLuc_01A.nam_psinh,'RRRR'),
                        TO_NUMBER(NVL(vc_PLuc_01A.so_psinh,0)),
                        TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_01,0)),
                        TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_02,0)),
                        TO_NUMBER(NVL(vc_PLuc_01A.lo_chuyen_03,0)),
                        '01');
            END IF;
        END LOOP;

        --Xu ly du lieu phu luc 01B
        OPEN c_Nam_ChuyenLo('0316');
        FETCH c_Nam_ChuyenLo INTO vc_Nam_ChuyenLo;
        CLOSE c_Nam_ChuyenLo;

        FOR vc_PLuc_01B IN c_PLuc_01B LOOP
            IF (vc_PLuc_01B.so_tt < 8 AND vc_PLuc_01B.so_tt > 1) THEN
                IF ((vc_PLuc_01B.so_psinh <> 0) OR ((vc_PLuc_01B.so_psinh = 0) AND(vc_PLuc_01B.lo_chuyen_01 <> 0 OR
                                                                                   vc_PLuc_01B.lo_chuyen_02 <> 0 OR
                                                                                   vc_PLuc_01B.lo_chuyen_03 <> 0 OR
                                                                                   vc_PLuc_01B.lo_chuyen_04 <> 0 OR
                                                                                   vc_PLuc_01B.lo_chuyen_05 <> 0 OR
                                                                                   vc_PLuc_01B.lo_chuyen_06 <> 0))) THEN
                    SELECT qlt_xltk_dtl_seq.NEXTVAL INTO v_PLuc01B_Id FROM dual;
                    INSERT INTO qlt_pluc_qtoan_tndn_01a (id,
                                                         quh_id,
                                                         quh_ltd,
                                                         nam_psinh,
                                                         lo_psinh,
                                                         loai)
                    VALUES (v_PLuc01B_Id,
                            v_Hdr_Id,
                            0,
                            TO_DATE(vc_PLuc_01B.nam_psinh,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.so_psinh,0)),
                            '02');
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_01 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_1,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_01,0)));
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_02 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_2,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_02,0)));
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_03 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_3,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_03,0)));
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_04 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_4,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_04,0)));
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_05 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_5,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_05,0)));
                END IF;

                IF (vc_PLuc_01B.lo_chuyen_06 <> 0) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_01ad (pdn1a_id,
                                                          nam_chuyen_lo,
                                                          lo_chuyen)
                    VALUES (v_PLuc01B_Id,
                            TO_DATE(vc_Nam_ChuyenLo.nam_chuyen_6,'RRRR'),
                            TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_06,0)));
                END IF;
            ELSE
                INSERT INTO qlt_pluc_qtoan_tndn_01b (id,
                                                     quh_id,
                                                     quh_ltd,
                                                     nam_psinh,
                                                     lo_psinh,
                                                     lo_chuyen_truoc,
                                                     lo_chuyen_tky,
                                                     lo_chuyen_sau,
                                                     loai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        TO_DATE(vc_PLuc_01B.nam_psinh,'RRRR'),
                        TO_NUMBER(NVL(vc_PLuc_01B.so_psinh,0)),
                        TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_01,0)),
                        TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_02,0)),
                        TO_NUMBER(NVL(vc_PLuc_01B.lo_chuyen_03,0)),
                        '02');
            END IF;
        END LOOP;

        --Xu ly du lieu phu luc tu 02 - 13
        FOR vc_PLuc_02_13 IN c_PLuc_02_13 LOOP
            IF (vc_PLuc_02_13.ky_hieu = 'A') THEN
                IF (vc_PLuc_02_13.so_dtnt IS NOT NULL) THEN
                    INSERT INTO qlt_pluc_qtoan_tndn_a(id,
                                                      quh_id,
                                                      quh_ltd,
                                                      dcp_id,
                                                      gia_tri,
                                                      kieu_dlieu,
                                                      so_tt)
                    VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                            v_Hdr_Id,
                            0,
                            vc_PLuc_02_13.dcp_id,
                            DECODE(vc_PLuc_02_13.kieu_dlieu,
                                   'D',DECODE(LENGTH(vc_PLuc_02_13.so_dtnt),4,'01/01/'||vc_PLuc_02_13.so_dtnt
                                                                           ,7,'01/'||vc_PLuc_02_13.so_dtnt
                                                                           ,vc_PLuc_02_13.so_dtnt),
                                   'C',DECODE(vc_PLuc_02_13.so_dtnt,'x','Y',vc_PLuc_02_13.so_dtnt),
                                   vc_PLuc_02_13.so_dtnt),
                            vc_PLuc_02_13.kieu_dlieu,
                            vc_PLuc_02_13.row_id);
                END IF;
            ELSE
                INSERT INTO qlt_pluc_qtoan_tndn_b(id,
                                                  quh_id,
                                                  quh_ltd,
                                                  dcp_id,
                                                  so_tt,
                                                  so_dtnt,
                                                  so_cqt,
                                                  ke_khai_sai,
                                                  stt_pluc)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        vc_PLuc_02_13.dcp_id,
                        vc_PLuc_02_13.so_tt,
                        TO_NUMBER(NVL(vc_PLuc_02_13.so_dtnt,0)),
                        TO_NUMBER(NVL(vc_PLuc_02_13.so_dtnt,0)),
                        NULL,
                        vc_PLuc_02_13.row_id);
            END IF;
        END LOOP;

        --Xu ly phu luc 14
        FOR vc_PLuc_14 IN c_PLuc_14 LOOP
            INSERT INTO qlt_pluc_qtoan_tndn_14 (id,
                                                quh_id,
                                                quh_ltd,
                                                ten_dia_chi,
                                                tnhap_nte,
                                                tnhap_vnd,
                                                thue_nte,
                                                thue_vnd,
                                                tnhap_tndn_nte,
                                                tnhap_tndn_vnd,
                                                tsuat_tndn,
                                                thue_tndn,
                                                thue_ktru)
            VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    vc_PLuc_14.ten_dia_chi,
                    TO_NUMBER(NVL(vc_PLuc_14.tnhap_nte,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.tnhap_vnd,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.thue_nte,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.thue_vnd,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.tnhap_tndn_nte,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.tnhap_tndn_vnd,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.tsuat_tndn,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.thue_tndn,0)),
                    TO_NUMBER(NVL(vc_PLuc_14.thue_ktru,0)));
        END LOOP;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);

        /*Ghi so ho so nhan*/
        v_Ds_Pluc := '107,';
        --Phu luc 1A, 1B
        OPEN c_PLuc_01A;
        FETCH c_PLuc_01A INTO vc_PLuc_01A;

        OPEN c_PLuc_01B;
        FETCH c_PLuc_01B INTO vc_PLuc_01B;

        IF c_PLuc_01A%FOUND OR c_PLuc_01B%FOUND THEN
            v_Ds_Pluc := v_Ds_Pluc||'108,';
        END IF;

        CLOSE c_PLuc_01A;
        CLOSE c_PLuc_01B;
        --Phu luc 2-> 13
        FOR vc_PLuc_Exist_02_13 IN c_PLuc_Exist_02_13 LOOP
            IF vc_PLuc_Exist_02_13.loai = '02' THEN
                v_Ds_Pluc := v_Ds_Pluc ||'109,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'03') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'110,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'04') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'111,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'05') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'112,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'06') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'113,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'07') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'114,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'08') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'115,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'09') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'116,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'10') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'117,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'11') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'118,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'12') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'119,';
            ELSIF INSTR(vc_PLuc_Exist_02_13.loai,'13') >0 THEN
                v_Ds_Pluc := v_Ds_Pluc ||'120,';
            END IF;
        END LOOP;
        --Kiem tra ton tai phu luc 14
        OPEN c_PLuc_14;
        FETCH c_PLuc_14 INTO vc_PLuc_14;
        IF c_PLuc_14%FOUND THEN
            v_Ds_Pluc := v_Ds_Pluc || '121,';
        END IF;
        CLOSE c_PLuc_14;

        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '05',
                         '03',
                         v_Ds_Pluc);
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_Qtoan_TNDN_Nam'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 17/03/2006
Muc dich: Thuc hien do du lieu to khai GTGT truc tiep vao CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai
*******************************************************************************/
    PROCEDURE Prc_TKhai_GTGT_TT(p_Record_Of_Header Record_Hdr,
                                p_Record_Dtnt Record_Dtnt,
                                p_TKhai_Exits_Id NUMBER,
                                p_Tthai_Tkhai VARCHAR2) IS

	    TYPE Record_Dtl_GTGT_TT IS RECORD(ctk_id NUMBER(10,0),
                                          so_tt NUMBER(3,0),
                                          so_dtnt NUMBER(20,2),
                                          so_cqt NUMBER(20,2),
                                          ke_khai_sai VARCHAR2(1));

        TYPE Array_Of_Record_Dtl IS TABLE OF Record_Dtl_GTGT_TT INDEX BY BINARY_INTEGER;
        v_Array_Of_Record_Dtl Array_Of_Record_Dtl;

        CURSOR c_TKhai_GTGT_TT IS
            SELECT *
            FROM rcv_v_tkhai_gtgt_tt
            WHERE (hdr_id = p_Record_Of_Header.id)
            ORDER BY so_tt;

        CURSOR c_Loai_Thue IS
            SELECT lte_ma_thue
            FROM qlt_dm_tkhai
            WHERE ma = p_Record_Of_Header.loai_tkhai;
        vc_Loai_Thue c_Loai_Thue%ROWTYPE;

        CURSOR c_So_Gtgt_Ktruoc IS
        	SELECT ABS(gtgt_tt.so_dtnt) so_ktruoc
        	FROM qlt_tkhai_gtgt_tt gtgt_tt
        		,qlt_tkhai_hdr hdr
        	WHERE (gtgt_tt.tkh_id = hdr.id)
              AND (gtgt_tt.tkh_ltd = hdr.ltd)
        	  AND (hdr.tin = p_Record_Of_Header.tin)
        	  AND (gtgt_tt.so_tt = '5')
        	  AND (gtgt_tt.so_dtnt < 0 )
        	  AND (hdr.ltd = 0)
        	  AND (hdr.kykk_tu_ngay = ADD_MONTHS(p_Record_Of_Header.kykk_tu_ngay, - 1));
        vc_So_Gtgt_Ktruoc c_So_Gtgt_Ktruoc%ROWTYPE;

        v_HanNop DATE;
        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_Ma_Thue VARCHAR(2);
        v_Thue_PSinh NUMBER(20,2);
        v_So_GTGT_KTruoc NUMBER(20,2);
        v_Check BOOLEAN := TRUE;
        v_CTieu3_Dtnt NUMBER(20,2);
        v_CTieu4_Dtnt NUMBER(20,2);
        v_CTieu5_Cqt NUMBER(20,2);
        v_Count NUMBER(10,0) := 0;
        v_Sai_SoHoc VARCHAR2(1) := NULL;
    BEGIN
        --Luu cac gia tri cua to khai GTGT truc tiep ra mang
        FOR vc_TKhai_GTGT_TT IN c_TKhai_GTGT_TT LOOP
            v_Count := v_Count + 1;
            v_Array_Of_Record_Dtl(v_Count).ctk_id := vc_TKhai_GTGT_TT.ctk_id;
            v_Array_Of_Record_Dtl(v_Count).so_tt := vc_TKhai_GTGT_TT.so_tt;
            v_Array_Of_Record_Dtl(v_Count).so_dtnt := NVL(vc_TKhai_GTGT_TT.gia_tri,0);
            v_Array_Of_Record_Dtl(v_Count).so_cqt := NVL(vc_TKhai_GTGT_TT.gia_tri,0);
            v_Array_Of_Record_Dtl(v_Count).ke_khai_sai := NULL;
        END LOOP;

        --Lay so GTGT ky truoc chuyen sang cua DTNT
        --"GTGT am ky truoc chuyen sang" (chi tieu 0) chi co voi cac to khai mau moi tu sau ky lap bo 02/2004
        IF (p_Record_Of_Header.kylb_tu_ngay > TO_DATE('01/02/2004','DD/MM/RRRR')) THEN
            --"GTGT am ky truoc chuyen sang" khong chuyen sang nam sau
            IF (TO_CHAR(p_Record_Of_Header.kylb_tu_ngay,'DD/MM') <> '01/01') THEN
                OPEN c_So_Gtgt_Ktruoc;
                FETCH c_So_Gtgt_Ktruoc INTO vc_So_GTGT_KTruoc;
                IF (c_So_Gtgt_Ktruoc%FOUND) THEN
                    v_So_GTGT_KTruoc := vc_So_GTGT_KTruoc.so_ktruoc;
                ELSE
                    v_So_GTGT_KTruoc := 0;
                END IF;
                CLOSE c_So_Gtgt_Ktruoc;
            ELSE
            	v_So_GTGT_KTruoc := 0;
            END IF;
        ELSE
            v_So_GTGT_KTruoc := 0;
        END IF;

        --Kiem tra so GTGT am ky truoc chuyen qua cua DTNT voi so cua CQT
        v_Array_Of_Record_Dtl(1).so_cqt := v_So_GTGT_KTruoc;
        IF (v_Array_Of_Record_Dtl(1).so_cqt <> v_Array_Of_Record_Dtl(1).so_dtnt) THEN
            v_Array_Of_Record_Dtl(1).ke_khai_sai := 'Y';
            v_Sai_SoHoc := 'Y';
        END IF;

        --Kiem tra so GTGT phat sinh trong ky cua DTNT voi so cua CQT
        --5 = 4 - 3 - So GTGT am ky truoc chuyen qua
        v_CTieu5_Cqt := v_Array_Of_Record_Dtl(5).so_dtnt -
                        v_Array_Of_Record_Dtl(4).so_dtnt -
                        v_So_GTGT_KTruoc;

        v_Array_Of_Record_Dtl(6).so_cqt := v_CTieu5_Cqt;
        IF (v_Array_Of_Record_Dtl(6).so_cqt <> v_Array_Of_Record_Dtl(6).so_dtnt) THEN
            v_Array_Of_Record_Dtl(6).ke_khai_sai := 'Y';
            v_Sai_SoHoc := 'Y';
        ELSE
            v_Array_Of_Record_Dtl(6).ke_khai_sai := NULL;
            v_Sai_SoHoc := NULL;
        END IF;

        --Neu phat sinh thue GTGT phai nop thi luu lai
        IF (v_Array_Of_Record_Dtl(7).so_dtnt >= 0) THEN
            v_Thue_PSinh := v_Array_Of_Record_Dtl(7).so_dtnt;
        END IF;

        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            /*Neu da ton tai to khai trong ky ke khai*/
            Qlt_Pck_Gdich.Prc_Lay_Thamso(p_Tthai_Tkhai,p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_TKHAI_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);

            /*Thong tin Header*/
            UPDATE qlt_tkhai_hdr
               SET co_loi = v_Sai_SoHoc,
                   co_loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                   ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                   --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                   so_tt_tk = p_Record_Of_Header.so_tt_tk,
                   ngay_nop = p_Record_Of_Header.ngay_nop,
                   kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                   kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                   kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                   kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                   tthai = '4', --To khai thay the
                   ngay_cap_nhat = p_Record_Of_Header.ngay_cap_nhat,
                   nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat
            WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);

            /*Thong tin Detail*/
            FOR i IN 1..v_Count LOOP
                UPDATE qlt_tkhai_gtgt_tt
                   SET so_dtnt = v_Array_Of_Record_Dtl(i).so_dtnt,
                       so_cqt = v_Array_Of_Record_Dtl(i).so_cqt,
                       ke_khai_sai = v_Array_Of_Record_Dtl(i).ke_khai_sai
                WHERE (tkh_id = p_TKhai_Exits_Id)
                  AND (tkh_ltd = 0)
                  AND (ctk_id = v_Array_Of_Record_Dtl(i).ctk_id);
            END LOOP;

            /*Thong tin bang phat sinh*/
            UPDATE qlt_psinh_tkhai
            SET thue_psinh = v_Thue_PSinh
            WHERE (tkh_id = p_TKhai_Exits_Id)
              AND (tkh_ltd = 0);

        ELSE --To khai chua ton tai
            --Neu co an dinh
            IF (Fnc_AnDinh_Exits(p_Record_Of_Header.tin,
                                 p_Record_Of_Header.loai_tkhai,
                                 p_Record_Of_Header.kykk_tu_ngay,
                                 p_Record_Of_Header.kykk_den_ngay)) THEN
                v_Tthai_Tkhai := '3'; --To khai nop cham
            ELSE --Neu chua co an dinh
                v_HanNop := Fnc_Han_Nop(p_Record_Of_Header.loai_tkhai,
                                        'TK',
                                        p_Record_Of_Header.kykk_den_ngay);
                IF (p_Record_Of_Header.ngay_nop <= v_HanNop) THEN
                    --To khai dung han
                    v_Tthai_Tkhai := '1'; --To khai chinh thuc
                ELSE
                    --To khai khong dung han
                    v_Tthai_Tkhai := '3'; --To khai nop cham
                END IF;
            END IF;

            --Thuc hien xu ly to khai
            Qlt_Pck_Gdich.Prc_Lay_Thamso(v_Tthai_Tkhai,
                                         p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Xu ly voi du lieu to khai Header
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
                                      co_loi,
                                      co_loi_ddanh,
                                      ghi_chu_loi,
                                      so_hieu_tep,
                                      so_tt_tk,
                                      ngay_cap_nhat,
                                      nguoi_cap_nhat)
            VALUES(v_Hdr_Id,
                   0,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   p_Record_Dtnt.ma_cqt,
                   p_Record_Dtnt.ma_tinh,
                   p_Record_Dtnt.ma_huyen,
                   p_Record_Of_Header.dia_chi,
                   p_Record_Dtnt.ma_phong,
                   p_Record_Dtnt.ma_canbo,
                   p_Record_Of_Header.loai_tkhai,
                   p_Record_Of_Header.ngay_nop,
                   p_Record_Of_Header.kylb_tu_ngay,
                   p_Record_Of_Header.kylb_den_ngay,
                   p_Record_Of_Header.kykk_tu_ngay,
                   p_Record_Of_Header.kykk_den_ngay,
                   v_Tthai_Tkhai,
                   v_Sai_SoHoc,
                   p_Record_Of_Header.co_loi_ddanh,
                   p_Record_Of_Header.ghi_chu_loi,
                   NULL,  --p_Record_Of_Header.so_hieu_tep,
                   p_Record_Of_Header.so_tt_tk,
                   p_Record_Of_Header.ngay_cap_nhat,
                   p_Record_Of_Header.nguoi_cap_nhat);

            --Xu ly voi du lieu to khai Detail
            FOR i IN 1..v_Count LOOP
                INSERT INTO qlt_tkhai_gtgt_tt (id,
                                               tkh_id,
                                               tkh_ltd,
                                               ctk_id,
                                               so_tt,
                                               so_dtnt,
                                               so_cqt,
                                               ke_khai_sai)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        v_Array_Of_Record_Dtl(i).ctk_id,
                        v_Array_Of_Record_Dtl(i).so_tt,
                        v_Array_Of_Record_Dtl(i).so_dtnt,
                        v_Array_Of_Record_Dtl(i).so_cqt,
                        v_Array_Of_Record_Dtl(i).ke_khai_sai);
            END LOOP;

            --Tinh so khau tru cho thue GTGT khau tru
            Qlt_Pck_Tinh_So.Prc_Tinh_So_Ktru(TRUNC(p_Record_Of_Header.kylb_tu_ngay,'MONTH')
                                                  ,p_Record_Of_Header.tin
                                                  ,'014'
                                                  ,'01'
                                                  ,v_Ma_Thue);
            --Tinh so phat sinh
            OPEN c_Loai_Thue;
            FETCH c_Loai_Thue INTO vc_Loai_Thue;
            IF (c_Loai_Thue%FOUND) THEN
                v_Ma_Thue := vc_Loai_Thue.lte_ma_thue;
            END IF;
            CLOSE c_Loai_Thue;

            --Insert du lieu vao bang phat sinh
            INSERT INTO qlt_psinh_tkhai (id,
                                         tkh_id,
                                         tkh_ltd,
                                         ccg_ma_cap,
                                         ccg_ma_chuong,
                                         lkn_ma_loai,
                                         lkn_ma_khoan,
                                         tmt_ma_muc,
                                         tmt_ma_tmuc,
                                         tmt_ma_thue,
                                         thue_psinh)
            VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    p_Record_Dtnt.ma_cap,
                    p_Record_Dtnt.ma_chuong,
                    p_Record_Dtnt.ma_loai,
                    p_Record_Dtnt.ma_khoan,
                    '014',
                    '01',
                    v_Ma_Thue,
                    v_Thue_PSinh);
        END IF;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        /*Ghi so ho so nhan*/
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '02',--To khai GTGT TT
                         '02',--To khai nop thue
                          NULL);
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_GTGT_TT'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 17/03/2006
Muc dich: Thuc hien do du lieu to khai GTGT truc tiep vao CSDL TKN_TC
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Record_Dtnt: Bien ban ghi chua thong tin DTNT
        - p_TKhai_Exits_Id: Id cua to khai da ton tai
        - p_Tthai_Tkhai: Trang thai cua to khai da ton tai
*******************************************************************************/
    PROCEDURE Prc_TKhai_QToan_GTGT_TT(p_Record_Of_Header Record_Hdr,
                                      p_Record_Dtnt Record_Dtnt,
                                      p_TKhai_Exits_Id NUMBER,
                                      p_Tthai_Tkhai VARCHAR2) IS
        CURSOR c_QToan_GTGT_TT IS
            SELECT *
            FROM rcv_v_tkhai_qtoan_gtgt_tt
            WHERE (hdr_id = p_Record_Of_Header.id)
            ORDER BY so_tt;

        CURSOR c_Loai_Thue IS
            SELECT lte_ma_thue
            FROM qlt_dm_qtoan
            WHERE ma = p_Record_Of_Header.loai_tkhai;
        vc_Loai_Thue c_Loai_Thue%ROWTYPE;

        v_HanNop DATE;
        v_Tthai_Tkhai VARCHAR2(1);
        v_Hdr_Id NUMBER(10);
        v_Ma_Thue VARCHAR(2);
        v_So_KKhai_QToan NUMBER(20,2) := 0;
        v_So_KKhai_TKhai NUMBER(20,2) := 0;
        v_So_CLech NUMBER(20,2) := 0;
        v_Count NUMBER(10,0) := 0;
        v_Sai_SoHoc VARCHAR2(1) := NULL;
    BEGIN
        OPEN c_Loai_Thue;
        FETCH c_Loai_Thue INTO vc_Loai_Thue;
        IF (c_Loai_Thue%FOUND) THEN
            v_Ma_Thue := vc_Loai_Thue.lte_ma_thue;
        END IF;
        CLOSE c_Loai_Thue;

        v_So_KKhai_TKhai := Fnc_So_KKhai_GTGT_TT(p_Record_Of_Header.tin,
                                                 '014',
                                                 '01',
                                                 v_Ma_Thue,
                                                 p_Record_Of_Header.kykk_tu_ngay);
        IF (p_TKhai_Exits_Id IS NOT NULL) THEN
            /*Neu da ton tai to khai trong ky ke khai*/
            Qlt_Pck_Gdich.Prc_Lay_Thamso_QToan(p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            /*Thuc hien backup to khai, phu luc va cac chung tu lien quan*/
            Qlt_Pck_TKhai.Prc_Backup_TKhai('QLT_QTOAN_HDR',
                                           p_Record_Of_Header.loai_tkhai,
                                           p_TKhai_Exits_Id);

            /*Thong tin Header*/
            UPDATE qlt_qtoan_hdr
            SET ngay_nop = p_Record_Of_Header.ngay_nop,
                kylb_tu_ngay = p_Record_Of_Header.kylb_tu_ngay,
                kylb_den_ngay = p_Record_Of_Header.kylb_den_ngay,
                kykk_tu_ngay = p_Record_Of_Header.kykk_tu_ngay,
                kykk_den_ngay = p_Record_Of_Header.kykk_den_ngay,
                tthai = '4', --To khai thay the
                loi_ddanh = p_Record_Of_Header.co_loi_ddanh,
                ghi_chu_loi = p_Record_Of_Header.ghi_chu_loi,
                nguoi_cap_nhat = p_Record_Of_Header.nguoi_cap_nhat,
                --so_hieu_tep = p_Record_Of_Header.so_hieu_tep,
                so_tt_tk = p_Record_Of_Header.so_tt_tk
            WHERE (id = p_TKhai_Exits_Id)
              AND (ltd = 0);

            /*Thong tin Detail*/
            FOR vc_QToan_GTGT_TT IN c_QToan_GTGT_TT LOOP
                UPDATE qlt_qtoan_dtl
                   SET so_dtnt = NVL(vc_QToan_GTGT_TT.gia_tri,0),
                       so_cqt = NVL(DECODE(vc_QToan_GTGT_TT.so_tt,5,
                                           DECODE(vc_QToan_GTGT_TT.gia_tri*(-1),ABS(vc_QToan_GTGT_TT.gia_tri),0,vc_QToan_GTGT_TT.gia_tri),
                                           vc_QToan_GTGT_TT.gia_tri),0),
                       ke_khai_sai = NULL
                WHERE (quh_id = p_TKhai_Exits_Id)
                  AND (quh_ltd = 0)
                  AND (ctq_id = vc_QToan_GTGT_TT.ctq_id);
                --Luu so ke khai quyet toan
                IF (vc_QToan_GTGT_TT.so_tt = 6) THEN
                    v_So_KKhai_QToan := NVL(vc_QToan_GTGT_TT.gia_tri,0);
                END IF;
            END LOOP;
            --Tinh so chenh lech quyet toan
            v_So_CLech := v_So_KKhai_QToan - v_So_KKhai_TKhai;
            /*Thong tin bang phat sinh*/
            UPDATE qlt_psinh_qtoan
            SET so_kkhai_qtoan = v_So_KKhai_QToan,
                so_kkhai_tkhai = v_So_KKhai_TKhai,
                so_clech = v_So_CLech
            WHERE (quh_id = p_TKhai_Exits_Id)
              AND (quh_ltd = 0);

        ELSE --To khai chua ton tai
            v_Tthai_Tkhai := '1'; --To khai chinh thuc
            --Thuc hien xu ly to khai
            Qlt_Pck_Gdich.Prc_Lay_Thamso_QToan(p_Record_Of_Header.loai_tkhai);
            Qlt_Pck_Control.Prc_Gan_Tin(p_Record_Of_Header.tin);
            Qlt_Pck_Control.Prc_Reset_Log_Id;

            SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            --Xu ly voi du lieu quyet toan Header
            INSERT INTO qlt_qtoan_hdr(id,
                                      ltd,
                                      tin,
                                      ten_dtnt,
                                      cqt_ma_cqt,
                                      hun_ma_huyen,
                                      hun_ma_tinh,
                                      ma_phong,
                                      ma_can_bo,
                                      dia_chi,
                                      nganh_kdoanh,
                                      dien_thoai,
                                      fax,
                                      email,
                                      dqt_ma,
                                      ngay_nop,
                                      kylb_tu_ngay,
                                      kylb_den_ngay,
                                      kykk_tu_ngay,
                                      kykk_den_ngay,
                                      tthai,
                                      loi_so_hoc,
                                      loi_ddanh,
                                      ghi_chu_loi,
                                      nguoi_cap_nhat,
                                      so_hieu_tep,
                                      so_tt_tk)
            VALUES (v_Hdr_Id,
                    0,
                    p_Record_Of_Header.tin,
                    p_Record_Of_Header.ten_dtnt,
                    p_Record_Dtnt.ma_cqt,
                    p_Record_Dtnt.ma_huyen,
                    p_Record_Dtnt.ma_tinh,
                    p_Record_Dtnt.ma_phong,
                    p_Record_Dtnt.ma_canbo,
                    p_Record_Of_Header.dia_chi,
                    NULL,
                    p_Record_Dtnt.dien_thoai,
                    p_Record_Dtnt.fax,
                    p_Record_Dtnt.email,
                    p_Record_Of_Header.loai_tkhai,
                    p_Record_Of_Header.ngay_nop,
                    p_Record_Of_Header.kylb_tu_ngay,
                    p_Record_Of_Header.kylb_den_ngay,
                    p_Record_Of_Header.kykk_tu_ngay,
                    p_Record_Of_Header.kykk_den_ngay,
                    v_Tthai_Tkhai, --Trang thai to khai
                    NULL,
                    p_Record_Of_Header.co_loi_ddanh,
                    p_Record_Of_Header.ghi_chu_loi,
                    p_Record_Of_Header.nguoi_cap_nhat,
                    NULL,  --p_Record_Of_Header.so_hieu_tep,
                    p_Record_Of_Header.so_tt_tk);

            --Xu ly voi du lieu Detail quyet toan
            FOR vc_QToan_GTGT_TT IN c_QToan_GTGT_TT LOOP
                INSERT INTO qlt_qtoan_dtl (id,
                                           quh_id,
                                           quh_ltd,
                                           ctq_id,
                                           so_tt,
                                           so_dtnt,
                                           so_cqt)
                VALUES (qlt_xltk_dtl_seq.NEXTVAL,
                        v_Hdr_Id,
                        0,
                        vc_QToan_GTGT_TT.ctq_id,
                        vc_QToan_GTGT_TT.so_tt,
                        NVL(vc_QToan_GTGT_TT.gia_tri,0),
                        NVL(DECODE(vc_QToan_GTGT_TT.so_tt,5,
                                   DECODE(vc_QToan_GTGT_TT.gia_tri*(-1),ABS(vc_QToan_GTGT_TT.gia_tri),0,vc_QToan_GTGT_TT.gia_tri),
                                   vc_QToan_GTGT_TT.gia_tri),0));
                --Luu so ke khai quyet toan
                IF (vc_QToan_GTGT_TT.so_tt = 6) THEN
                    v_So_KKhai_QToan := NVL(vc_QToan_GTGT_TT.gia_tri,0);
                END IF;
            END LOOP;
            --Tinh so chenh lech quyet toan
            v_So_CLech := v_So_KKhai_QToan - v_So_KKhai_TKhai;
            --Insert du lieu vao bang phat sinh
            INSERT INTO qlt_psinh_qtoan (id,
                                         quh_id,
                                         quh_ltd,
                                         ccg_ma_cap,
                                         ccg_ma_chuong,
                                         lkn_ma_loai,
                                         lkn_ma_khoan,
                                         tmt_ma_muc,
                                         tmt_ma_tmuc,
                                         tmt_ma_thue,
                                         so_kkhai_qtoan,
                                         so_kkhai_tkhai,
                                         so_clech)
            VALUES (qlt_xltk_hdr_seq.NEXTVAL,
                    v_Hdr_Id,
                    0,
                    p_Record_Dtnt.ma_cap,
                    p_Record_Dtnt.ma_chuong,
                    p_Record_Dtnt.ma_loai,
                    p_Record_Dtnt.ma_khoan,
                    '014',
                    '01',
                    v_Ma_Thue,
                    v_So_KKhai_QToan,
                    v_So_KKhai_TKhai,
                    v_So_Clech);
        END IF;

        /*Sau khi dua thanh cong du lieu 1 to khai tu CSDL trung gian sang
         CSDL QLT, thuc hien cap nhat trang thai*/
        UPDATE rcv_tkhai_hdr
        SET da_nhan = 'Y' --Cap nhat da chuyen thanh cong
        WHERE (id = p_Record_Of_Header.id);
        /*Ghi so ho so nhan*/
        Prc_So_Nhan_Hoso(p_Record_Of_Header,
                         p_Record_Dtnt,
                         '02', --To khai quyet toan thue GTGT TT
                         '03', --To khai quyet toan thue
                         '125');
        EXCEPTION
        WHEN OTHERS THEN
            ROLLBACK;
            QLT_PCK_CONTROL.Prc_Err_Log('Rcv_Pck_Chuyen_Dlieu_QLT.Prc_TKhai_QTOAN_GTGT_TT'
                                        , FALSE
                                        , NULL);
    END;
/*******************************************************************************
Nguoi lap: Khainhg
Ngay lap: 19/05/2006
Noi dung: Thuc hien ghi so nhan ho so
Tham so:
        - p_Record_Of_Header: Bien ban ghi chua du lieu cua mot record
                              trong bang RCV_TKHAI_HDR
        - p_Loai_Tkhai: Loai to khai
        - p_Nhom: Nhom to khai (to khai thue, to khai quyet toan)
        - p_Ma_Pluc: Danh sach phu luc dinh kem
        - p_Phong_Xly: Phong xu ly
*******************************************************************************/
    PROCEDURE Prc_So_Nhan_Hoso(p_Record_Of_Header Record_Hdr,
                               p_Record_Dtnt Record_Dtnt,
                               p_Loai_Tkhai VARCHAR2,
                               p_Nhom VARCHAR2,
                               p_Ma_Pluc VARCHAR2) IS
        --Kiem Lay so ho so da ton tai
        CURSOR c_So_Hoso_Exist(v_Kykk DATE, p_Dhs_Ma VARCHAR2) IS
    		SELECT hso.id, hso.so_hoso
    		FROM qhs_so_hoso hso
               , qhs_dm_hoso dm
    		WHERE (dm.ma = hso.dhs_ma)
    		AND (dm.loai = '01')
    		AND (dm.nhm_nhom IN ('02','03'))
    		AND (hso.kykk_tu_ngay = v_Kykk)
    		AND (hso.dhs_ma = p_Dhs_Ma)
    		AND (hso.tin = p_Record_Of_Header.Tin);
        --Lay ma loai ho so
        CURSOR c_Lay_Ma_Hoso IS
            SELECT loai_hoso
            FROM qlt_map_hoso_tkhai
            WHERE (nhom = p_Nhom)
            AND (loai_tkhai = p_Loai_Tkhai);
        CURSOR c_Dia_Chi(p_Dia_Chi VARCHAR2) IS SELECT p_Dia_Chi ||','||
                                                       H.ten ||','||
                                                       T.ten
                                                FROM qlt_tinh T, qlt_huyen H
                                                WHERE (T.ma_tinh = H.ma_tinh)
                                                And (T.ma_tinh = p_Record_Dtnt.ma_tinh)
                                                AND (H.ma_huyen = p_Record_Dtnt.ma_huyen);

        vc_So_Hoso_Exist c_So_Hoso_Exist%ROWTYPE;
        v_Ngay_Nop DATE := SYSDATE;
        v_Tthai VARCHAR2(2);
        v_So_Hoso VARCHAR2(30);
        v_Hdr_Id NUMBER(10);
        v_Dhs_Ma VARCHAR2(5):= NULL;
        v_User_Name VARCHAR2(100):=Qlt_Pck_Control.v_User_Name;
        v_Dia_Chi VARCHAR2(100) := p_Record_Of_Header.dia_chi;
    BEGIN
        /*Lay ma ho so*/
        OPEN c_Lay_Ma_Hoso;
        FETCH c_Lay_Ma_Hoso INTO v_Dhs_Ma;
        CLOSE c_Lay_Ma_Hoso;
        /*Lay dia chi*/
        OPEN c_Dia_Chi(p_Record_Of_Header.dia_chi);
        FETCH c_Dia_Chi INTO v_Dia_Chi;
        CLOSE c_Dia_Chi;
        /*Lay so ho so nhan*/
        v_So_Hoso := Fnc_Get_So_Hoso(SYSDATE, v_Dhs_Ma);
        /*Kiem tra ton tai ho so*/
        OPEN c_So_Hoso_Exist(p_Record_Of_Header.Kykk_Tu_Ngay, v_Dhs_Ma);
        FETCH c_So_Hoso_Exist INTO vc_So_Hoso_Exist;
        /*Neu da ton tai*/
        IF c_So_Hoso_Exist%FOUND THEN
            --Trang thai thay the
            v_Tthai := '03';
            --Insert thay the ho so moi
            /*Chon ID moi cho record*/
        	SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            INSERT INTO qhs_so_hoso(id,
                                dhs_ma,
                                so_hoso,
                                tin,
                                ten,
                                dia_chi,
                                kykk_tu_ngay,
                                kykk_den_ngay,
                                ngay_nhan,
                                nguoi_nop,
                                ngay_nhap,
                                nguoi_nhap,
                                han_xuly,
                                ngay_hen,
                                phong_xly,
                                ngay_tra,
                                ngay_nop,
                                so_hieu_tep,
                                tthai_hoso,
                                so_hoso_bsung,
                                kq_hoso_hoan)
            VALUES(v_Hdr_Id,
                   v_Dhs_Ma,
                   v_So_Hoso,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   v_Dia_Chi,
                   p_Record_Of_Header.kykk_Tu_Ngay,
                   p_Record_Of_Header.kykk_Den_Ngay,
                   TRUNC(p_Record_Of_Header.ngay_cap_nhat),
                   NULL,
                   TRUNC(SYSDATE),
                   p_Record_Of_Header.nguoi_cap_nhat,
                   NULL,
                   NULL,
                   p_Record_Of_Header.phong_xly,
                   NULL,
                   p_Record_Of_Header.ngay_nop,
                   NULL, --p_Record_Of_Header.so_hieu_tep,
                   v_Tthai,
                   vc_So_Hoso_Exist.so_hoso, --So ho so bo xung.
                   '01'); --Defaul 01
            /*Insert bang phu luc ho so*/
            IF p_Ma_Pluc IS NOT NULL THEN
                FOR v_Ma_Pluc IN (SELECT ma FROM qhs_dm_phuluc
                                            WHERE INSTR(p_Ma_Pluc,ma)>0
                                            AND dhs_ma = v_Dhs_Ma) LOOP
                    INSERT INTO qhs_pluc_hso(id,
                                             hso_id,
                                             dpl_ma,
                                             ghi_chu)
                    VALUES(qlt_xltk_hdr_seq.NEXTVAL,
                           v_Hdr_Id,
                           v_Ma_Pluc.ma,
                           NULL);
                END LOOP;
            END IF;
        ELSE /*Neu chua ton tai*/
            --Gan trang thai chinh thuc
            v_Tthai := '01';
            /*Chon ID moi cho record*/
        	SELECT qlt_xltk_hdr_seq.NEXTVAL INTO v_Hdr_Id FROM dual;
            INSERT INTO qhs_so_hoso(id,
                                dhs_ma,
                                so_hoso,
                                tin,
                                ten,
                                dia_chi,
                                kykk_tu_ngay,
                                kykk_den_ngay,
                                ngay_nhan,
                                nguoi_nop,
                                ngay_nhap,
                                nguoi_nhap,
                                han_xuly,
                                ngay_hen,
                                phong_xly,
                                ngay_tra,
                                ngay_nop,
                                so_hieu_tep,
                                tthai_hoso,
                                so_hoso_bsung,
                                kq_hoso_hoan)
            VALUES(v_Hdr_Id,
                   v_Dhs_Ma,
                   v_So_Hoso,
                   p_Record_Of_Header.tin,
                   p_Record_Of_Header.ten_dtnt,
                   v_Dia_Chi,
                   p_Record_Of_Header.kykk_Tu_Ngay,
                   p_Record_Of_Header.kykk_Den_Ngay,
                   TRUNC(p_Record_Of_Header.ngay_cap_nhat),
                   NULL,
                   TRUNC(SYSDATE),
                   p_Record_Of_Header.nguoi_cap_nhat,
                   NULL,
                   NULL,
                   p_Record_Of_Header.phong_xly,
                   NULL,
                   p_Record_Of_Header.ngay_nop,
                   NULL, --p_Record_Of_Header.so_hieu_tep,
                   v_Tthai,
                   NULL,
                   '01');
            /*Insert bang phu luc ho so*/
            IF p_Ma_Pluc IS NOT NULL THEN
                FOR v_Ma_Pluc IN (SELECT ma FROM qhs_dm_phuluc
                                            WHERE INSTR(p_Ma_Pluc,ma)>0
                                            AND dhs_ma = v_Dhs_Ma) LOOP
                    INSERT INTO qhs_pluc_hso(id,
                                             hso_id,
                                             dpl_ma,
                                             ghi_chu)
                    VALUES(qlt_xltk_hdr_seq.NEXTVAL,
                           v_Hdr_Id,
                           v_Ma_Pluc.ma,
                           NULL);
                END LOOP;
            END IF;
        END IF;
        CLOSE c_So_Hoso_Exist;
        --Cap nhat bang tep ho so
        UPDATE qhs_tep_hoso
        SET so_hoso = p_Record_Of_Header.so_tt_tk
        WHERE so_hieu = p_Record_Of_Header.so_hieu_tep
        AND dhs_ma = v_dhs_ma;
    END;
/*******************************************************************************
Nguoi lap: Nguyen Ta Anh
Ngay lap: 16/11/2005
Muc dich: Thuc hien do du lieu to khai tu CSDL trung gian vao CSDL TKN_TC
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
                   tkhai.loai,
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
                   hdr.co_gtrinh_02c,
                   hdr.phong_xly
            FROM rcv_tkhai_hdr hdr,
                 rcv_map_tkhai tkhai
            WHERE (hdr.loai_tkhai = tkhai.ma_tkhai)
              AND (hdr.da_nhan IS NULL)
              AND (hdr.khoa_so IS NULL)
            ORDER BY hdr.id;
        vc_Insert_Header_TKhai c_Insert_Header_TKhai%ROWTYPE;
        vr_Record_Hdr Record_Hdr;
        vr_Record_Dtnt Record_Dtnt;
        v_TKhai_Exits_Id NUMBER(10);
        v_Tthai_Tkhai VARCHAR2(1);

        CURSOR c_Ky_Khoa_So IS
            SELECT MAX(kylb_tu_ngay)
            FROM qlt_sothue_lock
            WHERE (loai_so = 'ST1B');
        v_Ky_Khoa_So DATE := '01-jan-1980';
    BEGIN
        OPEN c_Ky_Khoa_So;
        FETCH c_Ky_Khoa_So INTO v_Ky_Khoa_So;
        CLOSE c_Ky_Khoa_So;

        FOR vc_Insert_Header_TKhai IN c_Insert_Header_TKhai LOOP
            Begin
                vr_Record_Hdr.id := vc_Insert_Header_TKhai.id;
                vr_Record_Hdr.tin := vc_Insert_Header_TKhai.tin;
                vr_Record_Hdr.ten_dtnt := vc_Insert_Header_TKhai.ten_dtnt;
                vr_Record_Hdr.dia_chi := vc_Insert_Header_TKhai.dia_chi;
                vr_Record_Hdr.loai_tkhai := vc_Insert_Header_TKhai.loai_tkhai;
                vr_Record_Hdr.loai := vc_Insert_Header_TKhai.loai;
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
                vr_Record_Hdr.phong_xly := vc_Insert_Header_TKhai.phong_xly;
                --Kiem tra da khoa so chua
                IF (v_Ky_Khoa_So < vr_Record_Hdr.kylb_tu_ngay) THEN
                    --Neu chua khoa so
                    --Lay thong tin DTNT
                    Prc_Thong_Tin_Dtnt(vr_Record_Hdr.tin, vr_Record_Dtnt);
                    --Kiem tra to khai da ton tai trong ky ke khai chua
                    v_TKhai_Exits_Id := Fnc_TKhai_Exits(vr_Record_Hdr.tin,
                                                        vr_Record_Hdr.loai_tkhai,
                                                        vr_Record_Hdr.kykk_tu_ngay,
                                                        vr_Record_Hdr.kykk_den_ngay,
                                                        v_Tthai_Tkhai,
                                                        vr_Record_Hdr.loai);
                    --Neu la to khai
                    IF (vr_Record_Hdr.loai = 'TK') THEN
                        IF (vr_Record_Hdr.loai_tkhai = '14') THEN
                            --To khai GTGT
                            Prc_TKhai_GTGT(vr_Record_Hdr,
                                           vr_Record_Dtnt,
                                           v_TKhai_Exits_Id,
                                           v_Tthai_Tkhai);

                        ELSIF (vr_Record_Hdr.loai_tkhai = '26') THEN
                            --To khai TNDN quy
                            Prc_TKhai_TNDN_Quy(vr_Record_Hdr,
                                               vr_Record_Dtnt,
                                               v_TKhai_Exits_Id,
                                               v_Tthai_Tkhai);

                        ELSIF (vr_Record_Hdr.loai_tkhai = '02') THEN
                            --To khai GTGT truc tiep
                            Prc_TKhai_GTGT_TT(vr_Record_Hdr,
                                              vr_Record_Dtnt,
                                              v_TKhai_Exits_Id,
                                              v_Tthai_Tkhai);

                        ELSIF (vr_Record_Hdr.loai_tkhai = '24') THEN
                            --To khai TNGUYEN
                            Prc_TKhai_TNGUYEN(vr_Record_Hdr,
                                              vr_Record_Dtnt,
                                              v_TKhai_Exits_Id,
                                              v_Tthai_Tkhai);

                            --To khai TTDB
                        ELSIF (vr_Record_Hdr.loai_tkhai = '25') THEN
                            Prc_TKhai_TTDB(vr_Record_Hdr,
                                           vr_Record_Dtnt,
                                           v_TKhai_Exits_Id,
                                           v_Tthai_Tkhai);

                        END IF;
                    --Neu la quyet toan
                    ELSIF (vr_Record_Hdr.loai = 'QT') THEN
                        IF (vr_Record_Hdr.loai_tkhai = '05') THEN
                            --To khai quyet toan TNDN nam
                            Prc_TKhai_QToan_TNDN_Nam(vr_Record_Hdr,
                                                     vr_Record_Dtnt,
                                                     v_TKhai_Exits_Id,
                                                     v_Tthai_Tkhai);

                        ELSIF (vr_Record_Hdr.loai_tkhai = '02') THEN
                            --To khai quyet toan GTGT truc tiep
                            Prc_TKhai_QToan_GTGT_TT(vr_Record_Hdr,
                                                    vr_Record_Dtnt,
                                                    v_TKhai_Exits_Id,
                                                    v_Tthai_Tkhai);

                        ELSIF (vr_Record_Hdr.loai_tkhai = '08') THEN
                            --To khai quyet toan TNGUYEN
                            Prc_TKhai_Qtoan_TNGUYEN(vr_Record_Hdr,
                                                    vr_Record_Dtnt,
                                                    v_TKhai_Exits_Id,
                                                    v_Tthai_Tkhai);

                        END IF;
                    END IF;
                ELSE
                    --Neu da khoa so
                    UPDATE rcv_tkhai_hdr
                    SET khoa_so = 'Y'
                    WHERE (id = vr_Record_Hdr.id);
                END IF;
                COMMIT;
            EXCEPTION
                --Bo qua to khai bi loi, xu ly to tiep theo
                WHEN OTHERS THEN
                    ROLLBACK;
            END;
        END LOOP;
    END;
END;
/


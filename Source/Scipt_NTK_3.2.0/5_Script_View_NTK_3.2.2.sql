--01A_TNDN_DK
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_01A_TNDN_DK (HDR_ID, CTK_ID, SO_TT, TEN_CTIEU, GIA_TRI, KY_HIEU_CTIEU) AS SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '01A_TNDN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '01A_TNDN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
--01A_1_TNDN_DK
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_PLUC_01A_1_TNDN_DK
    (
        HDR_ID,
        SO_TT,
        ROW_ID,
        MA_SO_THUE,
        TEN_NHA_THAU,
        TY_LE_PHAN_BO,
        SO_THUE_PHAT_SINH_PHAI_NOP,
        GHI_CHU
    ) AS
    (
        /* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01A_1_TNDN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    ) ;    
--01B_TNDN_DK
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_01B_TNDN_DK (HDR_ID, CTK_ID, SO_TT, TEN_CTIEU, GIA_TRI, KY_HIEU_CTIEU) AS SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '01B_TNDN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '01B_TNDN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
--01B_1_TNDN_DK
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_PLUC_01B_1_TNDN_DK
    (
        HDR_ID,
        SO_TT,
        ROW_ID,
        MA_SO_THUE,
        TEN_NHA_THAU,
        TY_LE_PHAN_BO,
        SO_THUE_PHAT_SINH_PHAI_NOP,
        GHI_CHU
    ) AS
    (
        /* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
        SELECT
            dtl.hdr_id,
            dtl.so_tt               so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01B_1_TNDN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    ) ;    
--01_TAIN_DK
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_01_TAIN_DK (HDR_ID, CTK_ID, SO_TT, TEN_CTIEU, GIA_TRI, KY_HIEU_CTIEU) AS SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '01_TAIN_DK' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '01_TAIN_DK' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
-- 01_1_TAIN_DK
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_PLUC_01_1_TAIN_DK
    (
        HDR_ID,
        SO_TT,
        ROW_ID,
        MA_SO_THUE,
        TEN_NHA_THAU,
        TY_LE_PHAN_BO,
        SO_THUE_PHAT_SINH_PHAI_NOP,
        GHI_CHU
    ) AS
    (
        /* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.TEN_NHA_THAU)     TEN_NHA_THAU,
            MAX (dtl.TY_LE_PHAN_BO)       TY_LE_PHAN_BO,
            MAX (dtl.SO_THUE_PHAT_SINH_PHAI_NOP)     SO_THUE_PHAT_SINH_PHAI_NOP,
            MAX (dtl.GHI_CHU)              GHI_CHU
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_THAU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) TY_LE_PHAN_BO,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) SO_THUE_PHAT_SINH_PHAI_NOP,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  GHI_CHU
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01_1_TAIN_DK' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    ) ;
    
-- 03_TD_TAIN
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_TKHAI_03_TD_TAIN
    (
        HDR_ID,
        SO_TT,
        ROW_ID,
        NHA_MAY_TD,
        MA_SO_THUE,
        SAN_LUONG,
        GIA_TINH_THUE,
        THUE_PHAT_SINH,
        THUE_MIEN_GIAM,
        THUE_PHAI_NOP
    ) AS
    (
        /* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.NHA_MAY_TD) NHA_MAY_TD,
            MAX (dtl.MA_SO_THUE) MA_SO_THUE,
            MAX (dtl.SAN_LUONG)     SAN_LUONG,
            MAX (dtl.GIA_TINH_THUE)       GIA_TINH_THUE,
            MAX (dtl.THUE_PHAT_SINH)     THUE_PHAT_SINH,
            MAX (dtl.THUE_MIEN_GIAM)              THUE_MIEN_GIAM,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) NHA_MAY_TD,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) SAN_LUONG,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL ) GIA_TINH_THUE,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL ) THUE_PHAT_SINH,
                    DECODE (gdien.cot_06, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_MIEN_GIAM,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    ) ;   
    -- 03_1_TD_TAIN
CREATE OR REPLACE VIEW
    QLT_NTK.RCV_V_PLUC_03_1_TD_TAIN
    (
        HDR_ID,
        SO_TT,
        ROW_ID,
        STT,
        CHI_TIEU,
        MA_SO_THUE,
        CQT_QUAN_LY,
		CQT_PARENT_ID,
        TY_LE_PHAN_BO,
        THUE_PHAI_NOP
    ) AS
    (
        /* Formatted on 2011/08/11 13:39 (Formatter Plus v4.8.7) */
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.STT) STT,
            MAX (dtl.CHI_TIEU)     CHI_TIEU,
            MAX (dtl.MA_SO_THUE)       MA_SO_THUE,
            MAX (dtl.CQT_QUAN_LY)     CQT_QUAN_LY,
			MAX (dtl.CQT_PARENT_ID)     CQT_PARENT_ID,
            MAX (dtl.TY_LE_PHAN_BO)              TY_LE_PHAN_BO,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) STT,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) CHI_TIEU,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_QUAN_LY,
					DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_PARENT_ID,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)  TY_LE_PHAN_BO,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '03_1_TD_TAIN' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
    ) ; 

--01_TD_GTGT
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_TKHAI_01_TD_GTGT (HDR_ID, CTK_ID, SO_TT, TEN_CTIEU, GIA_TRI, KY_HIEU_CTIEU) AS SELECT
    DTL.HDR_ID ,
    DTL.CTK_ID ,
    MAX(DTL.SO_TT)         SO_TT ,
    MAX(DTL.TEN_CTIEU)     TEN_CTIEU ,
    MAX(DTL.GIA_TRI)       GIA_TRI ,
    MAX(DTL.KY_HIEU_CTIEU) KY_HIEU_CTIEU
FROM
    (
        SELECT
            TKD.HDR_ID ,
            GDIEN.ID ,
            GDIEN.SO_TT ,
            TKD.ROW_ID ,
            GDIEN.MA_CTIEU                                                       CTK_ID ,
            GDIEN.TEN_CTIEU                                                      TEN_CTIEU ,
            REPLACE(DECODE(GDIEN.COT_01, TKD.KY_HIEU, TKD.GIA_TRI, NULL),'%','')   GIA_TRI ,
            DECODE(GDIEN.COT_01, TKD.KY_HIEU, '['||GDIEN.KY_HIEU_CTIEU||']', NULL) KY_HIEU_CTIEU
        FROM
            QLT_NTK.RCV_TKHAI_DTL TKD,
            (
                SELECT
                    GD.*,
                    CT.KY_HIEU,
                    CT.KY_HIEU_CTIEU
                FROM
                    QLT_NTK.RCV_GDIEN_TKHAI GD,
                    QLT_NTK.RCV_MAP_CTIEU CT
                WHERE
                    CT.GDN_ID (+) = GD.ID
                AND GD.LOAI_DLIEU = '01_TD_GTGT' ) GDIEN
        WHERE
            GDIEN.KY_HIEU = TKD.KY_HIEU (+)
        AND TKD.LOAI_DLIEU (+)= '01_TD_GTGT' ) DTL
GROUP BY
    DTL.HDR_ID,
    DTL.CTK_ID;
        -- 01_2_TD_GTGT
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_01_2_TD_GTGT AS
(
        
        SELECT
            dtl.hdr_id,
            dtl.so_tt                so_tt,
            dtl.row_id                 row_id,
            MAX (dtl.STT) STT,
            MAX (dtl.TEN_NHA_MAY)     TEN_NHA_MAY,
            MAX (dtl.MA_SO_THUE)       MA_SO_THUE,
            MAX (dtl.CQT_QUAN_LY)     CQT_QUAN_LY,
            MAX (dtl.CQT_PARENT_ID)     CQT_PARENT_ID,
            MAX (dtl.TY_LE_PHAN_BO)              TY_LE_PHAN_BO,
            MAX (dtl.THUE_PHAI_NOP)              THUE_PHAI_NOP
        FROM
            (
                SELECT
                    tkd.hdr_id,
                    tkd.row_id row_id,
                    gdien.ID,
                    gdien.so_tt                                            so_tt,
                    DECODE (gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, NULL ) STT,
                    DECODE (gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, NULL ) TEN_NHA_MAY,
                    DECODE (gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, NULL ) MA_SO_THUE,
                    DECODE (gdien.cot_07, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_QUAN_LY,
                    DECODE (gdien.cot_08, tkd.ky_hieu, tkd.gia_tri, NULL ) CQT_PARENT_ID,
                    DECODE (gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, NULL)  TY_LE_PHAN_BO,
                    DECODE (gdien.cot_05, tkd.ky_hieu, tkd.gia_tri, NULL)  THUE_PHAI_NOP
                FROM
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                WHERE
                    (
                        ctieu.gdn_id = gdien.ID)
                AND (
                        ctieu.ky_hieu = tkd.ky_hieu)
                AND (
                        tkd.loai_dlieu = '01_2_TD_GTGT' )
                AND tkd.loai_dlieu = gdien.loai_dlieu) dtl
        GROUP BY
            dtl.hdr_id,
            dtl.so_tt,
            dtl.row_id
);

--01_BCTL_DK
create or replace view QLT_NTK.rcv_v_tkhai_01_bctl_dk as
select
    tmp.hdr_id          hdr_id ,
    tmp.row_id          row_id ,
    tmp.ma_ctieu        ma_ctieu ,
    tmp.so_tt           so_tt ,
    tmp.DAU_THO         DAU_THO ,
    tmp.CONDENSATE      CONDENSATE ,
    tmp.KHI_THIEN_NHIEN KHI_THIEN_NHIEN ,
    tmp.GHI_CHU         GHI_CHU
from
    (
        select
            dtl.hdr_id               hdr_id ,
            dtl.row_id               row_id ,
            dtl.ma_ctieu             ma_ctieu ,
            dtl.so_tt                so_tt ,
            max(dtl.DAU_THO)         DAU_THO ,
            max(dtl.CONDENSATE)      CONDENSATE ,
            max(dtl.KHI_THIEN_NHIEN) KHI_THIEN_NHIEN ,
            max(dtl.GHI_CHU)         GHI_CHU
        from
            QLT_NTK.rcv_gdien_tkhai gd,
            (
                select
                    tkd.hdr_id,
                    nvl(tkd.row_id,0) row_id,
                    gdien.id,
                    gdien.so_tt,
                    gdien.ma_ctieu,
                    decode(gdien.cot_01, tkd.ky_hieu, tkd.gia_tri, null) DAU_THO,
                    decode(gdien.cot_02, tkd.ky_hieu, tkd.gia_tri, null) CONDENSATE,
                    decode(gdien.cot_03, tkd.ky_hieu, tkd.gia_tri, null) KHI_THIEN_NHIEN,
                    decode(gdien.cot_04, tkd.ky_hieu, tkd.gia_tri, null) GHI_CHU
                from
                    QLT_NTK.rcv_tkhai_dtl tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
                where
                    (
                        ctieu.gdn_id = gdien.id)
                and (
                        ctieu.ky_hieu = tkd.ky_hieu)
                and (
                        tkd.loai_dlieu = '01_BCTL_DK') ) dtl
        where
            (
                gd.loai_dlieu = '01_BCTL_DK')
        and (
                dtl.id = gd.id)
        group by
            dtl.hdr_id,
            dtl.row_id,
            dtl.so_tt,
            dtl.ma_ctieu ) tmp;

--KHBS_01A_TNDN_DK
CREATE OR REPLACE VIEW QLT_NTK.RCV_V_PLUC_KHBS13 AS
SELECT   dtl.hdr_id, dtl.so_tt so_tt, dtl.row_id row_id,
            MAX (dtl.ten_ctieu_dc) ten_ctieu_dc, MAX (dtl.ma_ctieu) ma_ctieu,
            MAX (dtl.so_da_kk) so_da_kk, MAX (dtl.so_dieu_chinh)
                                                                so_dieu_chinh,
            MAX (dtl.so_chenh_lech) so_chenh_lech
       FROM QLT_NTK.rcv_gdien_tkhai gd,
            (SELECT tkd.hdr_id, tkd.row_id row_id, gdien.ID,
                    gdien.so_tt so_tt,
                    DECODE (gdien.cot_01,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) ten_ctieu_dc,
                    DECODE (gdien.cot_02,
                            tkd.ky_hieu, '[' || tkd.gia_tri || ']',
                            NULL
                           ) ma_ctieu,
                    DECODE (gdien.cot_03,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_da_kk,
                    DECODE (gdien.cot_04,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_dieu_chinh,
                    DECODE (gdien.cot_05,
                            tkd.ky_hieu, tkd.gia_tri,
                            NULL
                           ) so_chenh_lech
               FROM QLT_NTK.RCV_TKHAI_DTL tkd,
                    QLT_NTK.rcv_gdien_tkhai gdien,
                    QLT_NTK.rcv_map_ctieu ctieu
              WHERE (ctieu.gdn_id = gdien.ID)
                AND (ctieu.ky_hieu = tkd.ky_hieu)
                AND (   tkd.loai_dlieu = 'KHBS_01A_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01B_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_03_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_TAIN13'
                     OR tkd.loai_dlieu = 'KHBS_01_TTDB13'
                     OR tkd.loai_dlieu = 'KHBS_03_NTNN13'
                     OR tkd.loai_dlieu = 'KHBS_05_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_01_BVMT13'
                     OR tkd.loai_dlieu = 'KHBS_04_GTGT13'
                     OR tkd.loai_dlieu = 'KHBS_02_TNDN13'
                     OR tkd.loai_dlieu = 'KHBS_01_PHXD13'
                      OR tkd.loai_dlieu = 'KHBS_01_NTNN13'
                       OR tkd.loai_dlieu = 'KHBS_01_TBVMT13'
                       OR tkd.loai_dlieu = 'KHBS_01A_TNDN_DK'
                    )) dtl
      WHERE (   gd.loai_dlieu = 'KHBS_01A_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01B_TNDN13'
             OR gd.loai_dlieu = 'KHBS_03_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_TAIN13'
             OR gd.loai_dlieu = 'KHBS_01_TTDB13'
             OR gd.loai_dlieu = 'KHBS_03_NTNN13'
             OR gd.loai_dlieu = 'KHBS_05_GTGT13'
             OR gd.loai_dlieu = 'KHBS_01_BVMT13'
             OR gd.loai_dlieu = 'KHBS_04_GTGT13'
             OR gd.loai_dlieu = 'KHBS_02_TNDN13'
             OR gd.loai_dlieu = 'KHBS_01_PHXD13'
			 OR gd.loai_dlieu = 'KHBS_01_NTNN13'
       OR gd.loai_dlieu = 'KHBS_01_TBVMT13'
       OR gd.loai_dlieu = 'KHBS_01A_TNDN_DK'
            )
        AND (dtl.ID = gd.ID)
   GROUP BY dtl.hdr_id, dtl.so_tt, dtl.row_id;
			
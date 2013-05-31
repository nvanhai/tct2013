-- Start of DDL Script for Table TKN_OWNER.RCV_TKHAI_GTGT_KTRU_TKTN
-- Generated 12-Dec-2005 11:57:36 from TKN_OWNER@TKN_92

CREATE TABLE rcv_tkhai_gtgt_ktru_tktn
    (id                             NUMBER(10,0) NOT NULL,
    id_tkhai_rcv_qlt               NUMBER(10,0),
    ma_dtnt                        VARCHAR2(14) NOT NULL,
    ky_kekhai                      VARCHAR2(7) NOT NULL,
    loai_tkhai                     VARCHAR2(7) NOT NULL,
    ky_lapbo                       VARCHAR2(7) NOT NULL,
    br_dt                          NUMBER(20,2),
    brgtgt_dt                      NUMBER(20,2),
    brgtgt_thue                    NUMBER(20,2),
    br0_dt                         NUMBER(20,2),
    br5_dt                         NUMBER(20,2),
    br5_thue                       NUMBER(20,2),
    br10_dt                        NUMBER(20,2),
    br10_thue                      NUMBER(20,2),
    br20_dt                        NUMBER(20,2),
    br20_thue                      NUMBER(20,2),
    mv_dt                          NUMBER(20,2),
    mvgtgt_thue                    NUMBER(20,2),
    ktru_thue                      NUMBER(20,2),
    pnop_thue                      NUMBER(20,2),
    kytruoc_thue                   NUMBER(20,2),
    nopthieu_thue                  NUMBER(20,2),
    nopthua_thue                   NUMBER(20,2),
    dnop_thue                      NUMBER(20,2),
    htra_thue                      NUMBER(20,2),
    pnop_th_thue                   NUMBER(20,2),
    ngay_nop                       DATE,
    ngay_nhap                      DATE,
    nnhap                          VARCHAR2(20),
    tt_ghiso                       VARCHAR2(1),
    gchu                           VARCHAR2(50),
    chua_ktru_kytruoc              NUMBER(20,2),
    mv_thue                        NUMBER(20,2),
    nhapkhau_dt                    NUMBER(20,2),
    nhapkhau_thue                  NUMBER(20,2),
    tscd_dt                        NUMBER(20,2),
    tscd_thue                      NUMBER(20,2),
    mvgtgt_dt                      NUMBER(20,2),
    dchinh_mv_tang_dt              NUMBER(20,2),
    dchinh_mv_tang_thue            NUMBER(20,2),
    dchinh_mv_giam_dt              NUMBER(20,2),
    dchinh_mv_giam_thue            NUMBER(20,2),
    tong_ktru_thue                 NUMBER(20,2),
    brkct_dt                       NUMBER(20,2),
    dchinh_br_tang_dt              NUMBER(20,2),
    dchinh_br_tang_thue            NUMBER(20,2),
    dchinh_br_giam_dt              NUMBER(20,2),
    dchinh_br_giam_thue            NUMBER(20,2),
    pnop_thue1                     NUMBER(20,2),
    ktru_thue_luyke                NUMBER(20,2),
    denghi_htra_thue               NUMBER(20,2),
    tong_ktru_kysau                NUMBER(20,2),
    dien_thoai                     VARCHAR2(20),
    fax                            VARCHAR2(20),
    email                          VARCHAR2(100),
    mv_trongnuoc_dt                NUMBER(20,2),
    mv_trongnuoc_thue              NUMBER(20,2),
    tongthue_mvgtgt                NUMBER(20,2),
    tongdt_brgtgt                  NUMBER(20,2),
    tongthue_brgtgt                NUMBER(20,2),
    br_thue                        NUMBER(20,2),
    tong_ktru_thue1                NUMBER(20,2),
    co_gtrinh_02a                  VARCHAR2(1),
    co_gtrinh_02b                  VARCHAR2(1),
    co_gtrinh_02c                  VARCHAR2(1),
    da_nhan                        VARCHAR2(1),
    co_loi_ddanh                   VARCHAR2(1),
    so_hieu_tep                    VARCHAR2(10),
    so_hieu_tkhai                  VARCHAR2(10),
    khong_psinh                    VARCHAR2(1))
/

-- Create synonym RCV_TKHAI_GTGT_KTRU_TKTN
CREATE PUBLIC SYNONYM rcv_tkhai_gtgt_ktru_tktn
  FOR rcv_tkhai_gtgt_ktru_tktn
/

-- Grants for Table
GRANT DELETE ON rcv_tkhai_gtgt_ktru_tktn TO tkn
/
GRANT INSERT ON rcv_tkhai_gtgt_ktru_tktn TO tkn
/
GRANT SELECT ON rcv_tkhai_gtgt_ktru_tktn TO tkn
/
GRANT UPDATE ON rcv_tkhai_gtgt_ktru_tktn TO tkn
/
GRANT SELECT ON rcv_tkhai_gtgt_ktru_tktn TO tkn_read
/



-- Constraints for RCV_TKHAI_GTGT_KTRU_TKTN

ALTER TABLE rcv_tkhai_gtgt_ktru_tktn
ADD CONSTRAINT rcv_tkhai_pk PRIMARY KEY (id)
USING INDEX
/


-- Comments for RCV_TKHAI_GTGT_KTRU_TKTN

COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br_dt IS 'T�ng doanh thu trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br_thue IS 'Thue GTGT hang hoa dv ban ra trong ky'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br0_dt IS 'Doanh so xu�t kh�u'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br10_dt IS 'Doanh s� ch�u thu� su�t 10%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br10_thue IS 'Thu� GTGT ��u ra 10%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br20_dt IS 'Doanh s� ch�u thu� su�t 20%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br20_thue IS 'Thu� GTGT ��u ra 20%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br5_dt IS 'Doanh s� ch�u thu� su�t 5%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br5_thue IS 'Thu� GTGT ��u ra 5%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brgtgt_dt IS 'Doanh so hang hoa ch� thu� GTGT'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brgtgt_thue IS 'T�ng s� thu� GTGT hang ho� b�n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brkct_dt IS 'Ha�ng ho�a, d�ch vu� ba�n ra kho�ng ch�u thue�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.chua_ktru_kytruoc IS 'Thu� GTGT c�n ���c kh�u tr� k� tr��c chuy�n sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02a IS 'B�ng gi�i tr�nh 02A'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02b IS 'B�ng gi�i tr�nh 02B'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02c IS 'B�ng gi�i tr�nh 02C'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_loi_ddanh IS 'L�i ��nh danh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.da_nhan IS 'C�p nh�t t� khai th�nh c�ng'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_giam_dt IS '�I�u ch�nh gi�m doanh thu b�n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_giam_thue IS '�I�u ch�nh gi�m thu� GTGT ��u ra �� k� khai k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_tang_dt IS '�I�u ch�nh t�ng doanh thu b�n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_tang_thue IS '�I�u ch�nh t�ng thu� GTGT ��u ra �� k� khai k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_giam_dt IS '�I�u ch�nh gi�m doanh thu mua v�o'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_giam_thue IS '�I�u ch�nh gi�m thu� GTGT �� ���c kh�u tr� c�c k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_tang_dt IS '�I�u ch�nh t�ng doanh thu mua v�o'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_tang_thue IS '�I�u ch�nh t�ng thu� GTGT �� ���c kh�u tr� c�c k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.denghi_htra_thue IS 'S� thu� GTGT �� ngh� ho�n k� n�y'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dien_thoai IS '�I�n tho�i c�a doanh nghi�p'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dnop_thue IS 'Thu� �� n�p trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.email IS 'Email c�a doanh nghi�p'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.fax IS 'Fax c�a doanh nghi�p'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.gchu IS 'Ghi ch� t�nh tr�ng ghi s� trong tr��ng h�p ghi v�o QLT kh�ng th�nh c�ng'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.htra_thue IS 'Thu� �� ho�n trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.khong_psinh IS 'Kh�ng ph�t sinh ho�t ��ng mua b�n trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ktru_thue IS 'Thu� GTGT ���c kh�u tr�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ktru_thue_luyke IS 'Thu� GTGT ch�a ���c kh�u tr� lu� k� ��n k� n�y'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ky_kekhai IS 'd�ng mm/yyyy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ky_lapbo IS 'd�ng mm/yyyy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.kytruoc_thue IS 'S� d� k� tr��c chuy�n sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.loai_tkhai IS '1 : ch�nh th�c, 2: thay th�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ma_dtnt IS 'M� s� thu�, n�u l� DN ph� thu�c th� c� g�ch ngang � gi�a'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_dt IS 'Doanh s� mua v�o'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_thue IS 'Doanh s� h�ng h�a d�ch v� mua v�o trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_trongnuoc_dt IS 'Doanh s� HHDV mua v�o trong n��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_trongnuoc_thue IS 'Thu� HHDV mua v�o trong n��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mvgtgt_dt IS 'H�ng h�a d�ch v� mua v�o ch�u thu� GTGT'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mvgtgt_thue IS 'Thu� GTGT mua v�o'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ngay_nhap IS 'Ng�y nh�p t� khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ngay_nop IS 'Ng�y n�p t� khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nhapkhau_dt IS 'Gi� tr� h�ng h�a nh�p kh�u'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nhapkhau_thue IS 'Thu� GTGT h�ng h�a nh�p kh�u'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nnhap IS 'Ng��i c�p nh�t'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nopthieu_thue IS 'C�n n� c�a k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nopthua_thue IS 'C�n ���c kh�u tr� c�a k� tr��c'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_th_thue IS 'C�n ph�I n�p cu�i k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_thue IS 'Thu� GTGT ph�t sinh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_thue1 IS 'Thu� GTGT ph�I n�p v�o ng�n s�ch trong k�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.so_hieu_tep IS 'S� hi�u t�p'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.so_hieu_tkhai IS 'S� hi�u t� khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_kysau IS 'Thu� GTGT c�n ���c kh�u tr� chuy�n k� sau'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_thue IS 'T�ng s� thu� GTGT ���c kh�u tr�'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_thue1 IS 'kh�ng t�nh thu� GTGT c�n ���c kh�u tr� k� tr��c chuy�n sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongdt_brgtgt IS 'T�ng doanh thu b�n  tra (g�m c� �I�u ch�nh t�ng, gi�m b�n ra c�c k� tr��c)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongthue_brgtgt IS 'T�ng thu� b�n  tra (g�m c� �I�u ch�nh t�ng, gi�m b�n ra c�c k� tr��c)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongthue_mvgtgt IS 'T�ng thu� GTGT mua v�o (g�m c� �I�u ch�nh t�ng,gi�m mua v�o c�c k� tr��c)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tscd_dt IS 'Gi� tr� t�I s�n c� ��nh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tscd_thue IS 'Thu� GTGT t�I s�n c� ��nh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tt_ghiso IS 'T�nh tr�ng ghi s�, m�c ��nh l� ''0'', N�u ghi v�o QLT th�nh c�ng th� ghi l� ''1'', ng��c l�i ghi l� ''2'''
/

-- End of DDL Script for Table TKN_OWNER.RCV_TKHAI_GTGT_KTRU_TKTN


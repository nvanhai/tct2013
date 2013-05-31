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

COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br_dt IS 'Tæng doanh thu trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br_thue IS 'Thue GTGT hang hoa dv ban ra trong ky'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br0_dt IS 'Doanh so xuÊt khÈu'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br10_dt IS 'Doanh sè chÞu thuÕ suÊt 10%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br10_thue IS 'ThuÕ GTGT ®Çu ra 10%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br20_dt IS 'Doanh sè chÞu thuÕ suÊt 20%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br20_thue IS 'ThuÕ GTGT ®Çu ra 20%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br5_dt IS 'Doanh sè chÞu thuÕ suÊt 5%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.br5_thue IS 'ThuÕ GTGT ®Çu ra 5%'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brgtgt_dt IS 'Doanh so hang hoa chô thuÕ GTGT'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brgtgt_thue IS 'Tæng sè thuÕ GTGT hang ho¸ b¸n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.brkct_dt IS 'Haøng hoùa, dòch vuï baùn ra khoâng chòu thueá'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.chua_ktru_kytruoc IS 'ThuÕ GTGT cßn ®­îc khÊu trõ kú tr­íc chuyÓn sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02a IS 'B¶ng gi¶i tr×nh 02A'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02b IS 'B¶ng gi¶i tr×nh 02B'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_gtrinh_02c IS 'B¶ng gi¶i tr×nh 02C'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.co_loi_ddanh IS 'Lçi ®Þnh danh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.da_nhan IS 'CËp nhËt tê khai thµnh c«ng'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_giam_dt IS '§IÒu chØnh gi¶m doanh thu b¸n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_giam_thue IS '§IÒu chØnh gi¶m thuÕ GTGT ®Çu ra ®· kª khai kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_tang_dt IS '§IÒu chØnh t¨ng doanh thu b¸n ra'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_br_tang_thue IS '§IÒu chØnh t¨ng thuÕ GTGT ®Çu ra ®· kª khai kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_giam_dt IS '§IÒu chØnh gi¶m doanh thu mua vµo'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_giam_thue IS '§IÒu chØnh gi¶m thuÕ GTGT ®· ®­îc khÊu trõ c¸c kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_tang_dt IS '§IÒu chØnh t¨ng doanh thu mua vµo'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dchinh_mv_tang_thue IS '§IÒu chØnh t¨ng thuÕ GTGT ®· ®­îc khÊu trõ c¸c kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.denghi_htra_thue IS 'Sè thuÕ GTGT ®Ò nghÞ hoµn kú nµy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dien_thoai IS '§IÖn tho¹i cña doanh nghiÖp'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.dnop_thue IS 'ThuÕ ®· nép trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.email IS 'Email cña doanh nghiÖp'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.fax IS 'Fax cña doanh nghiÖp'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.gchu IS 'Ghi chó t×nh tr¹ng ghi sæ trong tr­êng hîp ghi vµo QLT kh«ng thµnh c«ng'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.htra_thue IS 'ThuÕ ®· hoµn trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.khong_psinh IS 'Kh«ng ph¸t sinh ho¹t ®éng mua b¸n trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ktru_thue IS 'ThuÕ GTGT ®­îc khÊu trõ'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ktru_thue_luyke IS 'ThuÕ GTGT ch­a ®­îc khÊu trõ luü kÕ ®Õn kú nµy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ky_kekhai IS 'd¹ng mm/yyyy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ky_lapbo IS 'd¹ng mm/yyyy'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.kytruoc_thue IS 'Sè d­ kú tr­íc chuyÓn sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.loai_tkhai IS '1 : chÝnh thøc, 2: thay thÕ'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ma_dtnt IS 'M· sè thuÕ, nÕu lµ DN phô thuéc th× cã g¹ch ngang ë gi÷a'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_dt IS 'Doanh sè mua vµo'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_thue IS 'Doanh sè hµng hãa dÞch vô mua vµo trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_trongnuoc_dt IS 'Doanh sè HHDV mua vµo trong n­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mv_trongnuoc_thue IS 'ThuÕ HHDV mua vµo trong n­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mvgtgt_dt IS 'Hµng hãa dÞch vô mua vµo chÞu thuÕ GTGT'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.mvgtgt_thue IS 'ThuÕ GTGT mua vµo'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ngay_nhap IS 'Ngµy nhËp tê khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.ngay_nop IS 'Ngµy nép tê khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nhapkhau_dt IS 'Gi¸ trÞ hµng hãa nhËp khÈu'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nhapkhau_thue IS 'ThuÕ GTGT hµng hãa nhËp khÈu'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nnhap IS 'Ng­êi cËp nhËt'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nopthieu_thue IS 'Cßn nî cña kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.nopthua_thue IS 'Cßn ®­îc khÊu trõ cña kú tr­íc'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_th_thue IS 'Cßn ph¶I nép cuèi kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_thue IS 'ThuÕ GTGT ph¸t sinh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.pnop_thue1 IS 'ThuÕ GTGT ph¶I nép vµo ng©n s¸ch trong kú'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.so_hieu_tep IS 'Sè hiÖu tÖp'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.so_hieu_tkhai IS 'Sè hiÖu tê khai'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_kysau IS 'ThuÕ GTGT cßn ®­îc khÊu trõ chuyÓn kú sau'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_thue IS 'Tæng sè thuÕ GTGT ®­îc khÊu trõ'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tong_ktru_thue1 IS 'kh«ng tÝnh thuÕ GTGT cßn ®­îc khÊu trõ kú tr­íc chuyÓn sang'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongdt_brgtgt IS 'Tæng doanh thu b¸n  tra (gåm c¶ ®IÒu chØnh t¨ng, gi¶m b¸n ra c¸c kú tr­íc)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongthue_brgtgt IS 'Tæng thuÕ b¸n  tra (gåm c¶ ®IÒu chØnh t¨ng, gi¶m b¸n ra c¸c kú tr­íc)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tongthue_mvgtgt IS 'Tæng thuÕ GTGT mua vµo (gåm c¶ ®IÒu chØnh t¨ng,gi¶m mua vµo c¸c kú tr­íc)'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tscd_dt IS 'Gi¸ trÞ tµI s¶n cè ®Þnh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tscd_thue IS 'ThuÕ GTGT tµI s¶n cè ®Þnh'
/
COMMENT ON COLUMN rcv_tkhai_gtgt_ktru_tktn.tt_ghiso IS 'T×nh tr¹ng ghi sæ, mÆc ®Þnh lµ ''0'', NÕu ghi vµo QLT thµnh c«ng th× ghi lµ ''1'', ng­îc l¹i ghi lµ ''2'''
/

-- End of DDL Script for Table TKN_OWNER.RCV_TKHAI_GTGT_KTRU_TKTN


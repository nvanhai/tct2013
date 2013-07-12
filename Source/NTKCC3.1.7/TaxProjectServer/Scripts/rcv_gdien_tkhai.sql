-- Start of DDL Script for Table QLT_OWNER.RCV_GDIEN_TKHAI
-- Generated 12-Dec-2005 11:28:04 from QLT_OWNER@QLT_93

CREATE TABLE rcv_gdien_tkhai
    (id                             NUMBER(10,0) NOT NULL,
    ten_ctieu                      VARCHAR2(200) NOT NULL,
    cot_01                         VARCHAR2(10),
    cot_02                         VARCHAR2(10),
    cot_03                         VARCHAR2(10),
    cot_04                         VARCHAR2(10),
    cot_05                         VARCHAR2(10),
    cot_06                         VARCHAR2(10),
    cot_07                         VARCHAR2(10),
    cot_08                         VARCHAR2(10),
    cot_09                         VARCHAR2(10),
    cot_10                         VARCHAR2(10),
    so_tt                          NUMBER(3,0),
    loai_dlieu                     VARCHAR2(4),
    ma_ctieu                       VARCHAR2(3))
/

-- Create synonym RCV_GDIEN_TKHAI
CREATE PUBLIC SYNONYM rcv_gdien_tkhai
  FOR rcv_gdien_tkhai
/

-- Grants for Table
GRANT DELETE ON rcv_gdien_tkhai TO qlt
/
GRANT INSERT ON rcv_gdien_tkhai TO qlt
/
GRANT SELECT ON rcv_gdien_tkhai TO qlt
/
GRANT UPDATE ON rcv_gdien_tkhai TO qlt
/
GRANT SELECT ON rcv_gdien_tkhai TO qlt_read
/



-- Constraints for RCV_GDIEN_TKHAI

ALTER TABLE rcv_gdien_tkhai
ADD CONSTRAINT rcv_gtk_pk PRIMARY KEY (id)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_GDIEN_TKHAI


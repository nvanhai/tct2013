CREATE TABLE rcv_gdien_tkhai
    (id                             NUMBER(10,0) NOT NULL,
    ten_ctieu                      VARCHAR2(200),
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
    cot_11                         VARCHAR2(10),
    so_tt                          NUMBER(3,0),
    loai_dlieu                     VARCHAR2(4),
    ma_ctieu                       VARCHAR2(3))
/
ALTER TABLE rcv_gdien_tkhai
ADD CONSTRAINT rcv_gtk_pk PRIMARY KEY (id)
USING INDEX
/

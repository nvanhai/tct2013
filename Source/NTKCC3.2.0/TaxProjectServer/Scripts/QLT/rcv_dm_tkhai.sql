CREATE TABLE rcv_dm_tkhai
    (ma                             VARCHAR2(2) NOT NULL,
    ten                            VARCHAR2(60) NOT NULL,
    kieu_ky                        VARCHAR2(1))
/
ALTER TABLE rcv_dm_tkhai
ADD CONSTRAINT rcv_dtk_pk PRIMARY KEY (ma)
USING INDEX
/


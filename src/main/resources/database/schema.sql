CREATE TABLE IF NOT EXISTS files_table
(
    id    BIGINT PRIMARY KEY ,
    original_file_name  VARCHAR(255) NOT NULL ,
    storage_file_name  VARCHAR(255) NOT NULL ,
    year  VARCHAR(4) NOT NULL ,
    month  VARCHAR(2) NOT NULL ,
    day  VARCHAR(2) NOT NULL ,
    time  VARCHAR(5) NOT NULL ,
    is_actual BOOLEAN
    );
CREATE SEQUENCE IF NOT EXISTS hibernate_sequence START WITH 1 INCREMENT BY 1;
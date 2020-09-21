CREATE DATABASE pythondocgen;
use pythondocgen;



CREATE TABLE user_log (
  user_id varchar(45) DEFAULT NULL,
  organization_id varchar(45) DEFAULT NULL,
  file_name varchar(45) DEFAULT NULL,
  folder_id varchar(45) DEFAULT NULL,
  user_name varchar(45) DEFAULT NULL,
  generated_date datetime DEFAULT NULL
);


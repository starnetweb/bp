-- ─────────────────────────────────────────────────────────
--  Research Portal — Database Schema
--  Run once: mysql -u root -p < schema.sql
-- ─────────────────────────────────────────────────────────

CREATE DATABASE IF NOT EXISTS research_portal
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE research_portal;

CREATE TABLE IF NOT EXISTS topics (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    topic        VARCHAR(600)  NOT NULL,
    level        ENUM('undergraduate','postgraduate') NOT NULL DEFAULT 'undergraduate',
    chapters       VARCHAR(50)   NOT NULL DEFAULT '1-5'
                                 COMMENT 'e.g. "3", "3-5", "1,3,5", "1-5"',
    front_matter   VARCHAR(100)  NOT NULL DEFAULT 'declaration,dedication,acknowledgements'
                                 COMMENT 'Comma-separated optional sections to include',
    custom_toc          TEXT          NULL     COMMENT 'Optional user-supplied TOC text',
    custom_instructions TEXT          NULL     COMMENT 'Optional extra instructions for document generation',
    phone        VARCHAR(30)   NOT NULL,
    email        VARCHAR(255)  NOT NULL,
    status       ENUM('pending','approved','processing','completed','failed')
                              NOT NULL DEFAULT 'pending',
    job_id       VARCHAR(100)  DEFAULT NULL COMMENT 'Flask job_id returned on approval',
    submitted_at TIMESTAMP     NOT NULL DEFAULT CURRENT_TIMESTAMP,
    approved_at  TIMESTAMP     NULL DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

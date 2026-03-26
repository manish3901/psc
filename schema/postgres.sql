-- Coupon management schema (PostgreSQL)

DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_type WHERE typname = 'coupon_name_status') THEN
    CREATE TYPE coupon_name_status AS ENUM ('Active', 'Disabled');
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_type WHERE typname = 'coupon_status') THEN
    CREATE TYPE coupon_status AS ENUM ('Active', 'Disabled');
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_type WHERE typname = 'coupon_validator_status') THEN
    CREATE TYPE coupon_validator_status AS ENUM ('Active', 'Disabled');
  END IF;
END
$$;

CREATE TABLE IF NOT EXISTS coupon_name (
  coupon_name_id SERIAL PRIMARY KEY,
  coupon_name VARCHAR(100) NOT NULL,
  description TEXT,
  status coupon_name_status DEFAULT 'Active',
  barcode_value VARCHAR(32)
);

CREATE UNIQUE INDEX IF NOT EXISTS ix_coupon_name_barcode_value ON coupon_name (barcode_value);

CREATE TABLE IF NOT EXISTS coupon_master (
  coupon_master_id SERIAL PRIMARY KEY,
  coupon_name_id INT NOT NULL REFERENCES coupon_name(coupon_name_id),
  coupon_type VARCHAR(50) NOT NULL,
  max_allowed INT DEFAULT 0,
  prize_level INT,
  status coupon_status DEFAULT 'Active',
  prize_image VARCHAR(255)
);

CREATE TABLE IF NOT EXISTS coupon_users (
  coupon_user_id SERIAL PRIMARY KEY,
  first_name VARCHAR(100) NOT NULL,
  last_name VARCHAR(100),
  mobile_no VARCHAR(10) NOT NULL,
  area_zone VARCHAR(100),
  coupon_master_id INT NOT NULL REFERENCES coupon_master(coupon_master_id),
  unique_code VARCHAR(5),
  created_at TIMESTAMP DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS ix_coupon_users_unique_code ON coupon_users (unique_code);
CREATE INDEX IF NOT EXISTS ix_coupon_users_mobile_no ON coupon_users (mobile_no);
CREATE INDEX IF NOT EXISTS ix_coupon_users_created_at ON coupon_users (created_at);

CREATE TABLE IF NOT EXISTS coupon_validators (
  id SERIAL PRIMARY KEY,
  mobile_no VARCHAR(10) UNIQUE NOT NULL,
  name VARCHAR(120),
  city VARCHAR(120),
  status coupon_validator_status DEFAULT 'Active',
  created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS ix_coupon_validators_mobile_no ON coupon_validators (mobile_no);


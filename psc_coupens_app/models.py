from datetime import datetime

from . import db


CouponNameStatus = db.Enum("Active", "Disabled", name="coupon_name_status", native_enum=False)
CouponStatus = db.Enum("Active", "Disabled", name="coupon_status", native_enum=False)
CouponValidatorStatus = db.Enum(
    "Active", "Disabled", name="coupon_validator_status", native_enum=False
)


class CouponName(db.Model):
    __tablename__ = "coupon_name"

    coupon_name_id = db.Column(db.Integer, primary_key=True)
    coupon_name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    status = db.Column(CouponNameStatus, default="Active", index=True)
    barcode_value = db.Column(db.String(32), unique=True, index=True)

    def __repr__(self):
        return f"<CouponName {self.coupon_name}>"


class CouponMaster(db.Model):
    __tablename__ = "coupon_master"

    coupon_master_id = db.Column(db.Integer, primary_key=True)
    coupon_name_id = db.Column(db.Integer, db.ForeignKey("coupon_name.coupon_name_id"), nullable=False)
    coupon_type = db.Column(db.String(50), nullable=False)
    max_allowed = db.Column(db.Integer, default=0)
    prize_level = db.Column(db.Integer)  # 1,2,3.. (lower = higher value)
    status = db.Column(CouponStatus, default="Active", index=True)
    prize_image = db.Column(db.String(255))  # 'uploads/coupons/<file>'

    coupon_name = db.relationship("CouponName", backref=db.backref("masters", lazy=True))

    def __repr__(self):
        return f"<CouponMaster {self.coupon_type} (Level {self.prize_level})>"


class CouponUser(db.Model):
    __tablename__ = "coupon_users"

    coupon_user_id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100))
    mobile_no = db.Column(db.String(10), nullable=False, index=True)
    area_zone = db.Column(db.String(100))
    coupon_master_id = db.Column(db.Integer, db.ForeignKey("coupon_master.coupon_master_id"), nullable=False)
    unique_code = db.Column(db.String(5), unique=True, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)

    coupon_master = db.relationship("CouponMaster", backref=db.backref("users", lazy=True))

    def __repr__(self):
        return f"<CouponUser {self.first_name} - {self.unique_code}>"


class CouponValidator(db.Model):
    __tablename__ = "coupon_validators"

    id = db.Column(db.Integer, primary_key=True)
    mobile_no = db.Column(db.String(10), unique=True, nullable=False, index=True)
    name = db.Column(db.String(120))
    city = db.Column(db.String(120))
    status = db.Column(CouponValidatorStatus, default="Active", index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)

    def __repr__(self):
        return f"<CouponValidator {self.mobile_no}>"


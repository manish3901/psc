import os
import secrets
import string
from pathlib import Path

from flask import Blueprint, flash, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

from . import db
from .models import CouponMaster, CouponName, CouponUser, CouponValidator

main = Blueprint("main", __name__)


def normalize_mobile(value):
    digits = "".join(ch for ch in (value or "") if ch.isdigit())
    return digits[-10:] if len(digits) >= 10 else digits


def generate_coupon_barcode_value(prefix="CN"):
    for _ in range(20):
        candidate = f"{prefix}{secrets.randbelow(10**10):010d}"
        if not CouponName.query.filter_by(barcode_value=candidate).first():
            return candidate
    return f"{prefix}{secrets.token_hex(6).upper()}"


def require_admin():
    return bool(session.get("is_admin"))


@main.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if require_admin():
        return redirect(url_for("main.admin_dashboard"))

    if request.method == "POST":
        password = request.form.get("password") or ""
        expected = os.getenv("ADMIN_PASSWORD") or ""
        if not expected:
            flash("ADMIN_PASSWORD is not set on the server.", "danger")
            return redirect(url_for("main.admin_login"))
        if secrets.compare_digest(password, expected):
            session["is_admin"] = True
            flash("Logged in.", "success")
            return redirect(url_for("main.admin_dashboard"))
        flash("Invalid password.", "danger")

    return render_template("admin/login.html")


@main.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    flash("Logged out.", "info")
    return redirect(url_for("main.admin_login"))


@main.route("/admin")
def admin_dashboard():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    coupon_names = CouponName.query.order_by(CouponName.coupon_name_id.desc()).all()
    coupon_masters = CouponMaster.query.order_by(CouponMaster.coupon_master_id.desc()).all()
    coupon_users = CouponUser.query.order_by(CouponUser.coupon_user_id.desc()).limit(500).all()
    coupon_validators = CouponValidator.query.order_by(CouponValidator.id.desc()).limit(500).all()

    return render_template(
        "admin/dashboard.html",
        coupon_names=coupon_names,
        coupon_masters=coupon_masters,
        coupon_users=coupon_users,
        coupon_validators=coupon_validators,
    )


@main.route("/admin/coupon/name/create", methods=["POST"])
def coupon_name_create():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    name = (request.form.get("coupon_name") or "").strip()
    description = (request.form.get("description") or "").strip() or None
    status = request.form.get("status", "Active")
    status = status if status in ("Active", "Disabled") else "Active"

    if not name:
        flash("Coupon name is required.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    new_name = CouponName(
        coupon_name=name,
        description=description,
        status=status,
        barcode_value=generate_coupon_barcode_value(),
    )
    db.session.add(new_name)
    db.session.commit()
    flash(f"Coupon Name '{name}' created.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/name/update", methods=["POST"])
def coupon_name_update():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    coupon_name_id = request.form.get("coupon_name_id")
    name = (request.form.get("coupon_name") or "").strip()
    description = (request.form.get("description") or "").strip() or None
    status = request.form.get("status", "Active")
    status = status if status in ("Active", "Disabled") else "Active"

    cn = CouponName.query.get(coupon_name_id)
    if not cn:
        flash("Coupon Name not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if not name:
        flash("Coupon name is required.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    cn.coupon_name = name
    cn.description = description
    cn.status = status
    db.session.commit()
    flash("Coupon Name updated.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/name/delete/<int:id>")
def coupon_name_delete(id):
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    cn = CouponName.query.get(id)
    if not cn:
        flash("Coupon Name not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if CouponMaster.query.filter_by(coupon_name_id=id).first():
        flash("Cannot delete: masters exist. Disable or delete masters first.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.delete(cn)
    db.session.commit()
    flash("Coupon Name deleted.", "success")
    return redirect(url_for("main.admin_dashboard"))


def _save_prize_image(file_storage, name_id, coupon_type):
    if not file_storage or not file_storage.filename:
        return None

    allowed = {".png", ".jpg", ".jpeg", ".webp"}
    ext = Path(file_storage.filename).suffix.lower()
    if ext not in allowed:
        raise ValueError("Prize image must be PNG/JPG/WEBP.")

    upload_dir = Path(main.root_path) / "static" / "uploads" / "coupons"
    upload_dir.mkdir(parents=True, exist_ok=True)
    fname = secure_filename(file_storage.filename)
    fname = f"cm_{name_id}_{coupon_type}_{fname}".replace(" ", "_")
    file_storage.save(str(upload_dir / fname))
    return f"uploads/coupons/{fname}"


@main.route("/admin/coupon/master/create", methods=["POST"])
def coupon_master_create():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    name_id = request.form.get("coupon_name_id")
    coupon_type = (request.form.get("coupon_type") or "").strip()
    max_allowed = request.form.get("max_allowed")
    prize_level = request.form.get("prize_level")
    status = request.form.get("status", "Active")
    status = status if status in ("Active", "Disabled") else "Active"

    if not name_id or not coupon_type:
        flash("Coupon Name and Coupon Type are required.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    try:
        max_allowed_val = (
            int(max_allowed) if max_allowed is not None and str(max_allowed).strip() != "" else 0
        )
        prize_level_val = (
            int(prize_level) if prize_level is not None and str(prize_level).strip() != "" else None
        )
    except ValueError:
        flash("Max Allowed and Prize Level must be numbers.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    prize_image_path = None
    try:
        prize_image_path = _save_prize_image(request.files.get("prize_image"), name_id, coupon_type)
    except ValueError as e:
        flash(str(e), "danger")
        return redirect(url_for("main.admin_dashboard"))

    new_master = CouponMaster(
        coupon_name_id=int(name_id),
        coupon_type=coupon_type,
        max_allowed=max_allowed_val,
        prize_level=prize_level_val,
        status=status,
        prize_image=prize_image_path,
    )
    db.session.add(new_master)
    db.session.commit()
    flash("Coupon Master created.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/master/update", methods=["POST"])
def coupon_master_update():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    master_id = request.form.get("coupon_master_id")
    name_id = request.form.get("coupon_name_id")
    coupon_type = (request.form.get("coupon_type") or "").strip()
    max_allowed = request.form.get("max_allowed")
    prize_level = request.form.get("prize_level")
    status = request.form.get("status", "Active")
    status = status if status in ("Active", "Disabled") else "Active"

    cm = CouponMaster.query.get(master_id)
    if not cm:
        flash("Coupon Master not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    try:
        max_allowed_val = (
            int(max_allowed) if max_allowed is not None and str(max_allowed).strip() != "" else 0
        )
        prize_level_val = (
            int(prize_level) if prize_level is not None and str(prize_level).strip() != "" else None
        )
    except ValueError:
        flash("Max Allowed and Prize Level must be numbers.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    cm.coupon_name_id = int(name_id)
    cm.coupon_type = coupon_type
    cm.max_allowed = max_allowed_val
    cm.prize_level = prize_level_val
    cm.status = status

    try:
        new_path = _save_prize_image(request.files.get("prize_image"), name_id, coupon_type)
        if new_path:
            cm.prize_image = new_path
    except ValueError as e:
        flash(str(e), "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.commit()
    flash("Coupon Master updated.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/master/delete/<int:id>")
def coupon_master_delete(id):
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    cm = CouponMaster.query.get(id)
    if not cm:
        flash("Coupon Master not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if CouponUser.query.filter_by(coupon_master_id=id).first():
        flash("Cannot delete: users exist. Disable it instead.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.delete(cm)
    db.session.commit()
    flash("Coupon Master deleted.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/user/delete/<int:id>")
def coupon_user_delete(id):
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    cu = CouponUser.query.get(id)
    if not cu:
        flash("Coupon User entry not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.delete(cu)
    db.session.commit()
    flash("Coupon User entry deleted.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/validator/create", methods=["POST"])
def coupon_validator_create():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    mobile_no = normalize_mobile(request.form.get("mobile_no"))
    name = (request.form.get("name") or "").strip() or None
    city = (request.form.get("city") or "").strip() or None
    status = request.form.get("status", "Active")
    status = status if status in ("Active", "Disabled") else "Active"

    if len(mobile_no) != 10:
        flash("Please enter a valid 10-digit mobile number.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if CouponValidator.query.filter_by(mobile_no=mobile_no).first():
        flash("This mobile is already in validation list.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.add(CouponValidator(mobile_no=mobile_no, name=name, city=city, status=status))
    db.session.commit()
    flash("Validation mobile added.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/validator/delete/<int:id>")
def coupon_validator_delete(id):
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    cv = CouponValidator.query.get(id)
    if not cv:
        flash("Validation entry not found.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    db.session.delete(cv)
    db.session.commit()
    flash("Validation entry deleted.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/coupon/validator/upload", methods=["POST"])
def coupon_validator_upload():
    if not require_admin():
        return redirect(url_for("main.admin_login"))

    excel = request.files.get("excel_file")
    update_existing = request.form.get("update_existing") == "1"
    sheet_name = (request.form.get("sheet_name") or "").strip()

    if not excel or not excel.filename:
        flash("Please select an Excel file.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if not (excel.filename or "").lower().endswith(".xlsx"):
        flash("Only .xlsx files are supported.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    from io import BytesIO

    try:
        from openpyxl import load_workbook

        wb = load_workbook(BytesIO(excel.read()), data_only=True)
    except Exception:
        flash("Unable to read the Excel file.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            flash(f"Sheet '{sheet_name}' not found.", "danger")
            return redirect(url_for("main.admin_dashboard"))
        ws = wb[sheet_name]
    else:
        ws = wb.active

    headers = [(c.value or "") for c in ws[1]]
    norm = []
    import re

    for h in headers:
        s = str(h).strip().lower()
        s = re.sub(r"\\s+", " ", s)
        norm.append(s)

    def idx(*candidates):
        for c in candidates:
            c = c.lower()
            if c in norm:
                return norm.index(c)
        return None

    mobile_idx = idx(
        "mobile",
        "mobile no",
        "mobile number",
        "phone",
        "phone no",
        "phone number",
        "contact",
        "contact no",
    )
    name_idx = idx("name", "full name", "member name", "customer name")
    city_idx = idx("city", "location", "area", "zone", "area/zone", "area zone")
    status_idx = idx("status")

    if mobile_idx is None:
        flash("Excel must have a Mobile column in the first row.", "danger")
        return redirect(url_for("main.admin_dashboard"))

    inserted = 0
    updated = 0
    skipped_invalid = 0
    skipped_existing = 0
    seen = set()

    for row in ws.iter_rows(min_row=2):
        mobile_val = row[mobile_idx].value if mobile_idx < len(row) else None
        mobile = normalize_mobile(mobile_val)
        if len(mobile) != 10:
            skipped_invalid += 1
            continue
        if mobile in seen:
            continue
        seen.add(mobile)

        name = row[name_idx].value if name_idx is not None and name_idx < len(row) else None
        city = row[city_idx].value if city_idx is not None and city_idx < len(row) else None
        status = row[status_idx].value if status_idx is not None and status_idx < len(row) else None
        status = (str(status).strip().title() if status is not None and str(status).strip() else "Active")
        if status not in ("Active", "Disabled"):
            status = "Active"

        name = (str(name).strip() if name is not None and str(name).strip() else None)
        city = (str(city).strip() if city is not None and str(city).strip() else None)

        existing = CouponValidator.query.filter_by(mobile_no=mobile).first()
        if existing:
            if update_existing:
                existing.name = name
                existing.city = city
                existing.status = status
                updated += 1
            else:
                skipped_existing += 1
            continue

        db.session.add(CouponValidator(mobile_no=mobile, name=name, city=city, status=status))
        inserted += 1

    db.session.commit()
    flash(
        f"Import done (sheet: {ws.title}). Inserted: {inserted}, Updated: {updated}, "
        f"Skipped existing: {skipped_existing}, Invalid: {skipped_invalid}.",
        "success",
    )
    return redirect(url_for("main.admin_dashboard"))


@main.route("/coupon/validate-mobile", methods=["POST"])
def validate_mobile():
    data = request.get_json(silent=True) or {}
    mobile = normalize_mobile(data.get("mobile"))
    cv = CouponValidator.query.filter_by(mobile_no=mobile, status="Active").first()
    if cv:
        return {"found": True, "name": cv.name or "", "mobile": cv.mobile_no}
    return {"found": False}


@main.route("/coupon/entry")
def coupon_entry():
    cn_code = (request.args.get("cn") or "").strip()
    selected_coupon_name = None

    masters_query = CouponMaster.query.filter_by(status="Active")
    if cn_code:
        selected_coupon_name = CouponName.query.filter_by(barcode_value=cn_code).first()
        if selected_coupon_name and selected_coupon_name.status == "Active":
            masters_query = masters_query.filter_by(coupon_name_id=selected_coupon_name.coupon_name_id)
        else:
            selected_coupon_name = None
            masters_query = masters_query.filter_by(coupon_name_id=-1)

    active_coupons = masters_query.all()
    return render_template(
        "coupon/coupon_user_entry.html",
        coupons=active_coupons,
        selected_coupon_name=selected_coupon_name,
        cn_code=cn_code,
        company_profile=None,
        company_name=os.getenv("COMPANY_NAME", "Prakrutik SparshCare"),
    )


@main.route("/coupon/register", methods=["POST"])
def coupon_register():
    first_name = request.form.get("first_name")
    last_name = request.form.get("last_name")
    mobile_no = normalize_mobile(request.form.get("mobile_no"))
    area_zone = request.form.get("area_zone")
    cn_code = (request.form.get("cn_code") or "").strip()

    if len(mobile_no) != 10:
        return {"success": False, "message": "Please enter a valid 10-digit mobile number."}

    cv = CouponValidator.query.filter_by(mobile_no=mobile_no, status="Active").first()
    if not cv:
        return {"success": False, "message": "Please try again with a registered mobile number."}

    def pick_master_with_spacing(masters, coupon_name_id=None):
        if not masters:
            return None

        remaining_per_master = {}
        for m in masters:
            used = CouponUser.query.filter_by(coupon_master_id=m.coupon_master_id).count()
            max_allowed = int(m.max_allowed or 0)
            remaining = max_allowed - used if max_allowed > 0 else 10**9
            if remaining <= 0:
                continue
            remaining_per_master[m.coupon_master_id] = remaining

        if not remaining_per_master:
            return None

        recent_levels = []
        last_master_id = None
        if coupon_name_id:
            last_row = (
                db.session.query(CouponUser.coupon_master_id)
                .join(CouponMaster, CouponMaster.coupon_master_id == CouponUser.coupon_master_id)
                .filter(CouponMaster.coupon_name_id == coupon_name_id)
                .order_by(CouponUser.coupon_user_id.desc())
                .first()
            )
            last_master_id = last_row[0] if last_row else None

            recent_levels = [
                (row[0] or 0)
                for row in (
                    db.session.query(CouponMaster.prize_level)
                    .join(CouponUser, CouponUser.coupon_master_id == CouponMaster.coupon_master_id)
                    .filter(CouponMaster.coupon_name_id == coupon_name_id)
                    .order_by(CouponUser.coupon_user_id.desc())
                    .limit(30)
                    .all()
                )
            ]

        remaining_by_level = {}
        for m in masters:
            rem = remaining_per_master.get(m.coupon_master_id)
            if not rem:
                continue
            lvl = int(m.prize_level or 999)
            remaining_by_level[lvl] = remaining_by_level.get(lvl, 0) + min(rem, 10**6)

        level_gap = {}
        total_remaining = sum(remaining_by_level.values()) or 1
        for lvl, rem in remaining_by_level.items():
            avg_gap = int(round((total_remaining / max(1, rem)) - 1))
            avg_gap = max(0, min(avg_gap, 50))
            if lvl == 1:
                avg_gap = max(avg_gap, 3)
            elif lvl == 2:
                avg_gap = max(avg_gap, 1)
            level_gap[lvl] = avg_gap

        def level_allowed(lvl):
            gap = level_gap.get(lvl, 0)
            if gap <= 0:
                return True
            return lvl not in recent_levels[:gap]

        eligible = [
            m
            for m in masters
            if m.coupon_master_id in remaining_per_master and level_allowed(int(m.prize_level or 999))
        ]
        if not eligible:
            eligible = [m for m in masters if m.coupon_master_id in remaining_per_master]
        if not eligible:
            return None

        eligible_no_repeat = [m for m in eligible if m.coupon_master_id != last_master_id]
        if eligible_no_repeat:
            eligible = eligible_no_repeat

        weights = [max(1, min(remaining_per_master.get(m.coupon_master_id, 1), 10**6)) for m in eligible]
        import random

        return random.choices(eligible, weights=weights, k=1)[0]

    masters_query = CouponMaster.query.filter_by(status="Active")
    selected_coupon_name = None
    if cn_code:
        selected_coupon_name = CouponName.query.filter_by(barcode_value=cn_code).first()
        if not selected_coupon_name or selected_coupon_name.status != "Active":
            return {"success": False, "message": "Invalid/disabled coupon barcode. Please scan a valid coupon."}
        masters_query = masters_query.filter_by(coupon_name_id=selected_coupon_name.coupon_name_id)

    masters = masters_query.all()
    if not masters:
        return {"success": False, "message": "No coupons available right now. Please try again later."}

    coupon_name_id = None
    if cn_code:
        coupon_name_id = selected_coupon_name.coupon_name_id
        existing = (
            db.session.query(CouponUser.coupon_user_id)
            .join(CouponMaster, CouponMaster.coupon_master_id == CouponUser.coupon_master_id)
            .filter(CouponMaster.coupon_name_id == coupon_name_id)
            .filter(CouponUser.mobile_no == mobile_no)
            .first()
        )
        if existing:
            return {
                "success": False,
                "message": "You have already participated for this coupon. Please contact support if this is a mistake.",
            }

    master = None
    for _ in range(5):
        master = pick_master_with_spacing(masters, coupon_name_id=coupon_name_id)
        if not master:
            break
        if master.max_allowed and master.max_allowed > 0:
            used_count = CouponUser.query.filter_by(coupon_master_id=master.coupon_master_id).count()
            if used_count >= master.max_allowed:
                master = None
                continue
        break

    if not master:
        return {"success": False, "message": "All coupon types are exhausted right now. Please try again later."}

    new_user_coupon = CouponUser(
        first_name=first_name,
        last_name=last_name,
        mobile_no=mobile_no,
        area_zone=area_zone,
        coupon_master_id=master.coupon_master_id,
        unique_code=None,
    )
    db.session.add(new_user_coupon)
    db.session.commit()

    return {"success": True, "coupon_user_id": new_user_coupon.coupon_user_id}


@main.route("/coupon/reveal-code", methods=["POST"])
def coupon_reveal_code():
    data = request.get_json(silent=True) or {}
    coupon_user_id = data.get("coupon_user_id")

    cu = CouponUser.query.get(coupon_user_id)
    if not cu:
        return {"success": False, "message": "Invalid coupon request."}, 400

    master = CouponMaster.query.get(cu.coupon_master_id)
    if not master or master.status != "Active":
        return {"success": False, "message": "Coupon is not available right now."}, 400

    prize_image_url = None
    if getattr(master, "prize_image", None):
        prize_image_url = url_for("static", filename=master.prize_image)

    if cu.unique_code:
        return {
            "success": True,
            "unique_code": cu.unique_code,
            "coupon_type": master.coupon_type,
            "prize_level": master.prize_level,
            "prize_image_url": prize_image_url,
        }

    import random

    unique_code = "".join(random.choices(string.digits, k=5))
    while CouponUser.query.filter_by(unique_code=unique_code).first():
        unique_code = "".join(random.choices(string.digits, k=5))

    cu.unique_code = unique_code
    db.session.commit()

    return {
        "success": True,
        "unique_code": unique_code,
        "coupon_type": master.coupon_type,
        "prize_level": master.prize_level,
        "prize_image_url": prize_image_url,
    }


from flask import Blueprint, render_template, request, flash

import logic

bp = Blueprint("feature", __name__)

MAX_BYTES = 8 * 1024 * 1024


@bp.route("/")
def index():
    return render_template("index.html", result=None)


@bp.route("/analyze", methods=["POST"])
def analyze():
    result = None
    try:
        if request.form.get("demo"):
            result = logic.analyze_demo()
        else:
            f = request.files.get("file")
            if f is None or not f.filename:
                flash("Lütfen bir .xlsx dosyası seçin ya da örnek veriyle deneyin.", "error")
            else:
                data = f.read(MAX_BYTES + 1)
                if len(data) > MAX_BYTES:
                    flash("Dosya 8 MB sınırını aşıyor. Daha küçük bir dosya yükleyin.", "error")
                else:
                    result = logic.analyze_workbook(data, f.filename)
    except logic.AuditError as exc:
        flash(str(exc), "error")
    except Exception:
        flash("Dosya işlenirken beklenmeyen bir sorun oluştu. Dosyanın geçerli bir .xlsx olduğunu kontrol edin.", "error")
    return render_template("index.html", result=result)

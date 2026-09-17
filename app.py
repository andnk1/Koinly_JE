"""
app.py
======
Koinly -> QuickBooks JE web app.

Upload your Koinly transaction-history CSV(s) and (optionally) a GL account
mapping file and beginning/ending balances; get back one Excel workbook
(Reconciliation / JE by Month / Total JE tabs) plus the QuickBooks-import
QB_JE.csv, unchanged from the desktop script's format.

Stateless by design: nothing uploaded or generated is written to disk.
Uploads are processed entirely in memory, and the generated files are held
just long enough for you to click the two download buttons (a short-lived
in-memory cache, purged automatically) -- there's no history, no database,
and no saved run to come back to later.
"""
import os
import secrets
import time

from flask import Flask, request, render_template, send_file, abort, Response
from io import BytesIO

import koinly_engine as ke
import xlsx_report as xr
import mapping_template as mt

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024  # 60 MB of uploads, plenty for CSV/xlsx

# Short-lived, in-memory result cache: token -> {xlsx, csv, ts, summary}.
# This is what lets the results page offer two separate download buttons
# without re-uploading or re-processing -- NOT a persisted run history. Any
# entry older than RESULT_TTL_SECONDS is dropped the next time anyone hits
# the app (and everything vanishes on restart/redeploy, by design).
_RESULTS = {}
RESULT_TTL_SECONDS = 15 * 60


def _purge_old_results():
    cutoff = time.time() - RESULT_TTL_SECONDS
    for token in [t for t, v in _RESULTS.items() if v["ts"] < cutoff]:
        _RESULTS.pop(token, None)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/mapping-template.csv", methods=["GET"])
def mapping_template():
    csv_text = mt.build_csv()
    return Response(
        csv_text,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=Mapping_Template.csv"},
    )


@app.route("/run", methods=["POST"])
def run():
    _purge_old_results()

    txn_files = [f for f in request.files.getlist("txn_files") if f and f.filename]
    mapping_upload = request.files.get("mapping_file")
    beginning = (request.form.get("beginning") or "").strip()
    ending = (request.form.get("ending") or "").strip()

    if not txn_files:
        return render_template("index.html", error="Please choose at least one Koinly transaction history CSV.")

    try:
        files = [(f.filename, f.read()) for f in txn_files]
        mapping_file = None
        if mapping_upload is not None and mapping_upload.filename:
            mapping_file = (mapping_upload.filename, mapping_upload.read())

        result = ke.process(files, mapping_file, beginning or None, ending or None)
    except ke.EngineError as e:
        return render_template("index.html", error=str(e))
    except Exception as e:  # noqa: BLE001 -- surface anything unexpected as a plain message, not a stack trace
        return render_template("index.html", error=f"Couldn't process those files: {e}")

    xlsx_bytes = xr.build_workbook(result)
    csv_text = ke.rows_to_csv_string(ke.QB_HEADER, result["qb_rows"])

    token = secrets.token_urlsafe(16)
    _RESULTS[token] = {
        "xlsx": xlsx_bytes,
        "csv": csv_text.encode("utf-8"),
        "ts": time.time(),
    }

    blank_tag_review = sorted(
        (key, amounts) for key, amounts in result["blank_tag_review"].items()
    )

    return render_template(
        "results.html",
        token=token,
        result=result,
        blank_tag_review=blank_tag_review,
    )


@app.route("/download/<token>/xlsx", methods=["GET"])
def download_xlsx(token):
    _purge_old_results()
    entry = _RESULTS.get(token)
    if not entry:
        abort(410, "This report has expired -- please run the tool again.")
    return send_file(
        BytesIO(entry["xlsx"]),
        as_attachment=True,
        download_name="Koinly_JE_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download/<token>/csv", methods=["GET"])
def download_csv(token):
    _purge_old_results()
    entry = _RESULTS.get(token)
    if not entry:
        abort(410, "This report has expired -- please run the tool again.")
    return send_file(
        BytesIO(entry["csv"]),
        as_attachment=True,
        download_name="QB_JE.csv",
        mimetype="text/csv",
    )


@app.errorhandler(410)
def gone(e):
    return render_template("index.html", error=str(e.description)), 410


@app.errorhandler(413)
def too_large(e):
    return render_template("index.html", error="Those files are too large (60 MB limit)."), 413


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

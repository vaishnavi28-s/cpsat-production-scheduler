"""
Finishing compatibility rules for print job combination.

Rules determine whether two print jobs can run together on the same press,
based on their finishing operations (lamination, UV, foil, emboss, varnish).
"""

import re

SEP_RE = re.compile(r"[,\|\+\/]+")
PROCESS_INKS = {"c", "m", "y", "k", "cyan", "magenta", "yellow"}

LAM_EQUIV = {
    "GSP GLOSS OPP": {"GSP GLOSS OPP", "GLUEABLE STAMPABLE GLOSS OPP"},
    "GLUEABLE STAMPABLE GLOSS OPP": {"GSP GLOSS OPP", "GLUEABLE STAMPABLE GLOSS OPP"},
    "GSP MYLAR": {"GSP MYLAR", "GLUEABLE STAMPABLE MYLAR"},
    "GLUEABLE STAMPABLE MYLAR": {"GSP MYLAR", "GLUEABLE STAMPABLE MYLAR"},
}

ALLOWED_ADDONS = {
    "GLOSS FILM LAM": {"none", "emboss"},
    "MATTE FILM LAM": {"none", "spot_uv", "foil", "emboss",
                       "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "SCUFF-RESISTANT MATTE FILM LAM": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                                       "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "SANDY MATTE FILM LAM": {"none", "spot_uv", "emboss", "spot_uv+emboss"},
    "SOFT TOUCH FILM LAM": {"none", "spot_uv", "foil", "emboss",
                            "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "LAYFLAT GLOSS FILM LAM": {"none", "emboss"},
    "LAYFLAT MATTE FILM LAM": {"none", "spot_uv", "foil", "emboss",
                               "spot_uv+foil", "spot_uv+foil+emboss", "foil+emboss"},
    "SCUFF-RESISTANT MATTE LAYFLAT": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                                      "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "SOFT TOUCH LAYFLAT": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                           "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "GSP GLOSS OPP": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                      "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "GSP MATTE OPP": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                      "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "GLOSS MYLAR/POLYESTER": {"none", "emboss"},
    "MATTE MYLAR/POLYESTER": {"none", "spot_uv", "foil", "emboss",
                              "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
    "GSP MYLAR": {"none", "spot_uv", "gritty_uv", "foil", "emboss",
                  "spot_uv+foil", "spot_uv+emboss", "spot_uv+foil+emboss", "foil+emboss"},
}

ALLOW_UV_NO_OTHER = True
SPECIAL_TYPES = {"4/1", "5/0", "5/1", "6/0", "6/1", "4/4"}


def tokenize(s: str) -> list:
    parts = [re.sub(r"\s+", " ", t.strip().lower())
             for t in SEP_RE.split(s or "") if t.strip()]
    return parts


def canonical(s: str) -> str:
    toks = [t for t in tokenize(s) if t]
    return "+".join(sorted(toks))


def count_inks(s: str) -> int:
    return len(tokenize(s))


def type_label(job) -> str:
    n1 = count_inks(job.INKSS1)
    n2 = count_inks(job.INKSS2)
    combos = {
        (4, 0): "4/0", (4, 1): "4/1", (5, 0): "5/0",
        (5, 1): "5/1", (6, 0): "6/0", (6, 1): "6/1", (4, 4): "4/4",
    }
    return combos.get((n1, n2), "OTHER")


def signature(job) -> str:
    tl = type_label(job)
    return f"{tl}|{canonical(job.INKSS1)}|{canonical(job.INKSS2)}"


def ink_key(t: str) -> str:
    s = re.sub(r"\s+", " ", (t or "").strip().lower()).replace("pms", "").strip()
    m = re.search(r"\d+", s)
    num = m.group(0) if m else ""
    words = re.sub(r"\d+", " ", s)
    words = re.sub(r"\s+", " ", words).strip()
    parts = ([num] if num else []) + (sorted(words.split()) if words else [])
    return " ".join(parts).strip()


def is_process_ink(token: str) -> bool:
    t = re.sub(r"\s+", " ", (token or "").strip().lower())
    return t in PROCESS_INKS


def special_ink_set(job) -> set:
    inks = tokenize(job.INKSS1) + tokenize(job.INKSS2)
    return {ink_key(ink) for ink in inks if ink and not is_process_ink(ink)}


def has_finish_token(job, attr: str, needle: str) -> bool:
    return any(needle in t for t in tokenize(getattr(job, attr, "")))


def uv_overall(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "uv overall")


def uv_spot(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "uv spot") or any(
        needle in t
        for t in tokenize(job.FINISHINGOP)
        for needle in ["spot uv", "spot gloss", "spot matte silkscreen",
                       "spot gritty uv", "spot dimensional uv", "spot dimmensional uv"]
    )


def has_emboss(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "emboss") or any(
        n in t for t in tokenize(job.FINISHINGOP) for n in ["emboss", "deboss"]
    )


def has_foil(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "foil stamp") or any(
        "foil" in t for t in tokenize(job.FINISHINGOP)
    )


def press_varnish(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "press varnish") or any(
        n in t for t in tokenize(job.FINISHINGOP)
        for n in ["press varnish", "aqueous", "varnish"]
    )


def has_lamination(job) -> bool:
    return has_finish_token(job, "FINISHTYPE", "lamination") or any(
        n in t for t in tokenize(job.FINISHINGOP)
        for n in ["polypropylene", "layflat", "lay flat", "opp", "mylar", "polyester", "gsp"]
    )


def uv_overall_sheen(job):
    fo_tokens = tokenize(job.FINISHINGOP)
    overall = [t for t in fo_tokens if "uv" in t and "spot" not in t]
    if not overall:
        return None
    text = " | ".join(overall)
    if "gritty" in text:
        return "gritty"
    if "matte" in text:
        return "matte"
    if "gloss" in text:
        return "gloss"
    return None


def lam_pool(job) -> str:
    ops = " ".join(tokenize(job.FINISHINGOP))

    def has(*words):
        return all(w in ops for w in words)

    if has("gloss", "polypropylene"):                         return "GLOSS FILM LAM"
    if has("matte", "polypropylene") and not has("scuff"):    return "MATTE FILM LAM"
    if has("scuff", "matte", "polypropylene"):                return "SCUFF-RESISTANT MATTE FILM LAM"
    if "sandy" in ops:                                        return "SANDY MATTE FILM LAM"
    if has("soft", "touch", "poly"):                          return "SOFT TOUCH FILM LAM"
    if has("layflat", "gloss") or (has("lay", "flat") and "gloss" in ops): return "LAYFLAT GLOSS FILM LAM"
    if (has("layflat", "matte") or (has("lay", "flat") and "matte" in ops)) and not has("scuff"): return "LAYFLAT MATTE FILM LAM"
    if has("scuff", "layflat", "matte") or (has("lay", "flat") and "scuff" in ops and "matte" in ops): return "SCUFF-RESISTANT MATTE LAYFLAT"
    if has("soft", "touch", "layflat") or (has("lay", "flat") and "soft touch" in ops): return "SOFT TOUCH LAYFLAT"
    if has("gsp", "gloss", "opp") or has("glueable", "stampable", "gloss", "opp"): return "GSP GLOSS OPP"
    if has("gsp", "matte", "opp"): return "GSP MATTE OPP"
    if has("gloss", "mylar") or has("gloss", "polyester"): return "GLOSS MYLAR/POLYESTER"
    if has("matte", "mylar") or has("matte", "polyester"): return "MATTE MYLAR/POLYESTER"
    if has("glueable", "stampable", "mylar") or "gsp mylar" in ops: return "GSP MYLAR"
    return None


def addon_bucket(job) -> str:
    spot = uv_spot(job)
    gritty = any("gritty matte uv" in t for t in tokenize(job.FINISHINGOP)) and not uv_overall(job)
    foil = has_foil(job)
    emb = has_emboss(job)
    if gritty and not (spot or foil or emb):
        return "gritty_uv"
    key = []
    if spot:
        key.append("spot_uv")
    if foil:
        key.append("foil")
    if emb:
        key.append("emboss")
    return "none" if not key else "+".join(key)


def lam_equiv_set(pool: str) -> set:
    return LAM_EQUIV.get(pool, {pool})


def allowed_non_uv_partner(partner) -> bool:
    if has_lamination(partner) or press_varnish(partner) or has_foil(partner):
        return False
    if uv_spot(partner) or has_emboss(partner):
        return True
    return ALLOW_UV_NO_OTHER


def finishing_conflict(job_a, job_b) -> bool:
    """Returns True if two jobs cannot run together due to finishing incompatibility."""

    # Overall UV cannot combine with spot-only UV
    if (uv_overall(job_a) and uv_spot(job_b) and not uv_overall(job_b)) or \
       (uv_overall(job_b) and uv_spot(job_a) and not uv_overall(job_a)):
        return True

    auv, buv = uv_overall(job_a), uv_overall(job_b)
    if auv or buv:
        if (auv and (has_lamination(job_b) or press_varnish(job_b))) or \
           (buv and (has_lamination(job_a) or press_varnish(job_a))):
            return True
        if (auv and has_foil(job_b)) or (buv and has_foil(job_a)):
            return True
        if auv and buv:
            sa, sb = uv_overall_sheen(job_a), uv_overall_sheen(job_b)
            if sa is None or sb is None or sa != sb:
                return True
        if auv and not allowed_non_uv_partner(job_b):
            return True
        if buv and not allowed_non_uv_partner(job_a):
            return True

    alam, blam = has_lamination(job_a), has_lamination(job_b)
    if alam or blam:
        if (alam and press_varnish(job_b)) or (blam and press_varnish(job_a)):
            return True
        if alam and blam:
            pa, pb = lam_pool(job_a), lam_pool(job_b)
            if not pa or not pb:
                return True
            if pb not in lam_equiv_set(pa):
                return True
            adda, addb = addon_bucket(job_a), addon_bucket(job_b)
            if adda not in ALLOWED_ADDONS.get(pa, set()) or \
               addb not in ALLOWED_ADDONS.get(pb, set()):
                return True
            return False
        lam_job = job_a if alam else job_b
        other = job_b if alam else job_a
        pool = lam_pool(lam_job)
        if not pool:
            return True
        if addon_bucket(other) not in ALLOWED_ADDONS.get(pool, set()):
            return True

    if press_varnish(job_a) and (uv_overall(job_b) or has_lamination(job_b)):
        return True
    if press_varnish(job_b) and (uv_overall(job_a) or has_lamination(job_a)):
        return True

    return False

"""Spectral Technique Extrapolation Lab (STE Lab).

Legacy package name inside Ratio_extrapol. 1.5.1: any compiled path containing
tasto is excluded before the PHIL/MCGILL ordinario fallback (nested
Philharmonia arco-sul-tasto must not be ingested as ordinario). 1.5.0: Media
stays mean(IOWA, Orchidea). Extra collections enrich that Media only through
L = ln(effect / ordinario) when anchors overlap (transfer_field.mode=pooled).
Absolute Phil/McGill levels never enter Media. mode=single reproduces 1.4.0 L.
"""

__version__ = "1.5.1"

__all__ = ["__version__"]

"""Builds Stem_Cell_Microeconomics.xlsx and app/scenario_data.js from ONE set of parameter tables.

The spreadsheet is the source of truth for the calculator's literature-informed scenarios:
  CM  = team thesis: iPSC-derived cardiomyocytes (heart muscle cells)
  HSC = cord-blood hematopoietic stem cell (CD34+) expansion
Edit the add(...) rows below, rerun `py build_spreadsheet.py`, and both outputs update.
All values are USD, price year 2026 unless the source note says otherwise (no inflation adjustment applied).
"""
import json
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
# Sources. verification: READ = full page opened during research (24 Oct 2026);
# SNIPPET = value seen only in a search-result summary; verify before citing in the paper.
# (id, title, year, url, verification, system, side, finding, direction, linked inputs)
# ---------------------------------------------------------------------------
SOURCES = [
    # --- iPSC / cardiomyocyte (team thesis) ---
    ("C1", "Kempf H et al. Controlling Expansion and Cardiomyogenic Differentiation of Human Pluripotent Stem Cells in Scalable Suspension Culture. Stem Cell Reports", 2014,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC4264033/", "READ", "Human iPSC/ESC, stirred-tank bioreactor (Earth, 3D)", "Earth",
     "Bioreactor CM purity: HES3 68.6% +/- 8.7%; iPSC line 58% +/- 30.3% (runs 27-88%); ~40 million CMs per run; all 3 runs from batch-expanded cells failed.",
     "supportive (Earth baseline; shows variability)", "q_E, p_E"),
    ("C2", "Halloin C et al. Continuous WNT Control Enables Advanced hPSC Cardiac Processing and Prognostic Surface Marker Identification in Chemically Defined Suspension Culture. Stem Cell Reports", 2019,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC6700605/", "READ", "Human PSC, suspension bioreactor (Earth, 3D)", "Earth",
     "93.5% +/- 5.0% cTnT+ at 150 ml (n=3); 89.1% +/- 7.8% at 350-500 ml; up to ~4-4.5e8 CMs per run.", "supportive (Earth best case)", "q_E, g_E"),
    ("C3", "Prondzynski M et al. Efficient and reproducible generation of human iPSC-derived cardiomyocytes in stirred suspension (bioRxiv 2024; Nat Commun 2024)", 2024,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC10925150/", "READ", "Human iPSC, 19 bioreactor runs, 10 lines (Earth, 3D)", "Earth",
     ">90% cTnT+; ~2.5 CMs per input iPSC; ~$569 materials per run of ~124M CMs (research grade, excl. labor) vs ~$1,075 monolayer.",
     "supportive (Earth baseline)", "g_E, q_E, E_consumables"),
    ("C4", "Chemically defined cardiomyocyte differentiation protocol (STAR Protocols)", 2020,
     "https://www.cell.com/star-protocols/fulltext/S2666-1667(20)30002-2", "SNIPPET", "Human iPSC, hundreds of batches, >50 lines", "Earth",
     "Successful differentiation (>85% purity) in over 80% of runs.", "supportive (run success rate)", "p_E"),
    ("C5", "Kim JH, Kawase E, Bharti K, Karnieli O, Arakawa Y, Stacey G. Perspectives on the cost of goods for hPSC banks for manufacture of cell therapies. npj Regenerative Medicine 7:54", 2022,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC9522845/", "READ", "GMP hPSC banking", "Earth",
     "300-vial bank ~$300,000 (~$1,000/vial; range $1,000-10,000 across banks); labor ~50% of cost; outsourced testing ~$10k/lot; facility upkeep ~$2.2M/yr.",
     "supportive (cost)", "E_cells, E_labor, E_facility, E_qc"),
    ("C6", "Jha R, Wu Q, Singh M et al. Simulated Microgravity and 3D Culture Enhance Induction, Viability, Proliferation and Differentiation of Cardiac Progenitors from Human Pluripotent Stem Cells. Scientific Reports 6:30956", 2016,
     "https://www.nature.com/articles/srep30956", "SNIPPET", "Human PSC, random positioning machine vs 3D at 1g", "Space",
     "~99% CM purity, ~90% viability, 1.5-4-fold more CMs per input stem cell vs 3D culture at normal gravity.", "supportive (simulated microgravity)", "r_g, r_q"),
    ("C7", "Rampoldi A et al. Space microgravity improves proliferation of human iPSC-derived cardiomyocytes. Stem Cell Reports 17:2272", 2022,
     "https://www.cell.com/stem-cell-reports/fulltext/S2213-6711(22)00416-7", "SNIPPET", "Human iPSC cardiac progenitors, ISS 3 weeks", "Space",
     "Spheres ~3x larger, ~20x more nuclei than 1g controls; higher proliferation; better calcium handling. No purity comparison seen.", "supportive (real spaceflight)", "r_g"),
    ("C8", "Wnorowski A, Sharma A et al. Effects of Spaceflight on Human Induced Pluripotent Stem Cell-Derived Cardiomyocyte Structure and Function. Stem Cell Reports", 2019,
     "https://www.cell.com/stem-cell-reports/fulltext/S2213-6711(19)30367-4", "SNIPPET", "Human iPSC-CMs, ISS ~5.5 weeks", "Space",
     "2,635 differentially expressed genes; altered calcium handling; structure unchanged; reverted after return. Says nothing about yield.", "contextual", "s"),
    ("C9", "Timilsina S et al. Enhanced self-renewal of human pluripotent stem cells by simulated microgravity. npj Microgravity 8:22", 2022,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC9253108/", "READ", "Human PSC, 2D clinostat vs adherent 1g (no 3D control)", "Space",
     "Pluripotency genes up >2-fold, more cells, less induced differentiation. Supports 'stays pluripotent'; may also slow directed differentiation.", "mixed", "r_g, r_q"),
    ("C10", "Blaber EA et al. Microgravity Reduces the Differentiation and Regenerative Potential of Embryonic Stem Cells. Stem Cells and Development 24:2605", 2015,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC4652210/", "SNIPPET", "Mouse ESC, 15 days actual spaceflight", "Space",
     "Differentiation inhibited across germ layers; stemness retained.", "mixed (supports pluripotency, against differentiation)", "r_g, r_q"),
    ("C11", "Heartseed HS-001 dosing (Nikon press release)", 2025,
     "https://www.nikon.com/company/news/2025/0203_01/", "SNIPPET", "iPSC-CM clinical product", "Earth",
     "50 million (low) and 150 million (high) cardiomyocytes per patient.", "contextual (dose size)", "none (context)"),
    ("C12", "Ncardia large-scale hiPSC-cardiomyocyte manufacturing (Axion application note)", "not extracted",
     "https://www.axionbiosystems.com/resources/application-note/controlled-large-scale-manufacturing-hipsc-derived-cardiomyocytes", "SNIPPET", "Industrial iPSC-CM batch", "Earth",
     "7.5e9 cells at 87% cTnT+ in one batch.", "supportive (scale)", "N0, q_E"),
    ("C13", "Nanobioreactor detection of space-associated hematopoietic stem and progenitor cell aging. Cell Stem Cell (Jamieson lab)", 2025,
     "https://www.cell.com/cell-stem-cell/fulltext/S1934-5909(25)00270-X", "SNIPPET", "Human HSPC, ISS 32-45 days", "Space",
     "Aging signatures and ~5x more mutations than ground controls: a radiation/stress risk for any stem-cell product.", "contradictory (risk)", "p_S, r_q"),
    ("C14", "Phipson B et al. Evaluation of variability in human kidney organoids. Nature Methods 16:79", 2019,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC6634992/", "SNIPPET", "Human iPSC kidney organoids", "Earth",
     "Batch (not clone) is the main source of variation. Kidney differentiation is less controlled than cardiac: a stronger case for a quality lever.", "contextual (kidney)", "none (context)"),
    # --- Projections used by the space viability case ---
    ("P1", "SpaceOrbitals: How Cheap Is Starship V3 Really? A Plain-English Look at the 2026 $/kg-to-LEO Math", 2026,
     "https://spaceorbitals.com/articles/starship-v3-cost-per-kg-2026/", "READ", "Launch price projection (secondary source)", "Space",
     "Starship target '<$200/kg' (SpaceX since 2019); Musk May 2026: '$100-200 per kg' fully reusable (aspirational); '$500-1,000 (realistic mid-decade)'. No published Starship price.",
     "projection", "L"),
    ("P2", "Payload Space: Varda Conducts First Reentry Mission into Australia", 2025,
     "https://payloadspace.com/varda-conducts-first-reentry-mission-into-australia/", "READ", "Free-flyer reentry capsule (Varda)", "Space",
     "AFRL's Capt. Ashwin Rao: Varda 'reduces the cost of a reentry flight test from on the order of $100M to around $5M.' Capsule capacity not stated.",
     "projection", "C_return, S_facility"),
    # --- HSC scenario ---
    ("S1", "Wagner JE et al. Phase I/II Trial of StemRegenin-1 Expanded Umbilical Cord Blood Hematopoietic Stem Cells Supports Testing as a Stand-Alone Graft. Cell Stem Cell", 2016,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC4881386/", "READ", "Human cord blood CD34+, GMP clinical", "Earth",
     "CD34+ median 330-fold expansion (range 67-848); TNC 854-fold; 15-day culture. 2 of 20 products failed lot release; later 13 all passed.",
     "supportive (Earth baseline)", "g_E, p_E"),
    ("S2", "Horwitz ME et al. Omidubicel vs standard myeloablative umbilical cord blood transplantation: phase 3 randomized study. Blood 138(16):1429", 2021,
     "https://ashpublications.org/blood/article/138/16/1429/476235/", "SNIPPET", "Human cord blood CD34+, approved product (Omisirge)", "Earth",
     "CD34+ median 130-fold (32-233); ~6.6e8 CD34+ per product.", "supportive (Earth baseline)", "g_E, N0"),
    ("S3", "Single UM171 expanded cord blood transplant is feasible and safe. Experimental Hematology", 2017,
     "https://www.exphem.org/article/S0301-472X(17)30299-0/fulltext", "SNIPPET", "Human cord blood CD34+", "Earth",
     "CD34+ median net 35-fold expansion (n=7).", "supportive (Earth baseline)", "g_E"),
    ("S4", "Cohen S et al. Hematopoietic stem cell transplantation using single UM171-expanded cord blood: phase 1-2 study. Lancet Haematology", 2019,
     "https://www.thelancet.com/journals/lanhae/article/PIIS2352-3026(19)30202-9/abstract", "SNIPPET", "Human cord blood CD34+", "Earth",
     "26 of 27 units expanded successfully (96%).", "supportive (Earth baseline)", "p_E"),
    ("S5", "Harrison RP, Zylberberg E, Ellison S, Levine BL. CAR-T cell therapy manufacturing: modelling the effect of offshore production on aggregate cost of goods. Cytotherapy 21(2):224-233", 2019,
     "https://www.sciencedirect.com/science/article/pii/S1465324919300039", "SNIPPET", "CAR-T (different cell product)", "Earth",
     "Autologous COGS ~$95,780 per dose; allogeneic ~$4,460 per dose (2019 USD, not inflation-adjusted).", "contextual (other system)", "E_other, S_other"),
    ("S6", "ten Ham RMT et al. What does cell therapy manufacturing cost? A framework and methodology. Cytotherapy 22(7):388-397", 2020,
     "https://www.isct-cytotherapy.org/article/S1465-3249(20)30496-5/fulltext", "SNIPPET", "Academic cell-therapy facilities (EU)", "Earth",
     "Costing framework (fixed vs variable per-batch costs, 2018 EUR, 5% failure assumption). Per-batch totals not extracted.", "contextual (method)", "E_* cost structure"),
    ("S7", "HRSA Advisory Council on Blood Stem Cell Transplantation: Access to transplantation (Champlin R., slides)", "not extracted",
     "https://bloodstemcell.hrsa.gov/sites/default/files/bloodstemcell/about/advisory-council/meetings/access-to-transplantation.pdf", "SNIPPET", "Unrelated cord blood unit procurement", "Earth",
     "Cord blood unit procurement ~$30,000-$60,000+.", "supportive (cost)", "E_cells, S_cells"),
    ("S8", "UK National CAR T Clinical Panel: manufacturing failure analysis. Blood Cancer Journal", 2025,
     "https://www.nature.com/articles/s41408-025-01225-9", "SNIPPET", "CAR-T (different cell product)", "Earth",
     "LBCL manufacturing failure 3.87% (38/981).", "contextual (other system)", "p_E"),
    ("S9", "NASA ISS Updated Commercial and Marketing Pricing Policy (summary on eoPortal)", 2021,
     "https://www.eoportal.org/ftp/satellite-missions/iss/ISS-Pricing-08032021/ISS-Pricing.html", "READ", "ISS commercial access", "Space",
     "2021 full-value pricing: upmass $20,000/kg, downmass $40,000/kg, crew time $130,000/hour (2019 policy: $3,000/kg, $6,000/kg, $17,500/hr).",
     "supportive (cost)", "L, crew_rate, down_rate"),
    ("S10", "Payload Space: SpaceX Increases Rideshare Prices", 2025,
     "https://payloadspace.com/spacex-increases-rideshare-prices/", "SNIPPET", "Falcon 9 Transporter rideshare to SSO", "Space",
     "Rideshare $6,500/kg (2025), $7,000/kg (2026). Free-flyer orbit, not ISS delivery.", "contextual (cost floor)", "L"),
    ("S11", "NASA Office of Inspector General IG-18-016: NASA's Management of Commercial Resupply Services Contracts", 2018,
     "https://oig.nasa.gov/docs/IG-18-016.pdf", "SNIPPET", "ISS commercial cargo (CRS)", "Space",
     "Estimated SpaceX cargo cost ~$63,200/kg (CRS-1), ~$71,800/kg projected (CRS-2), as reported by secondary sources.", "supportive (cost ceiling)", "L"),
    ("S12", "Wang P et al. Spaceflight/microgravity inhibits the proliferation of hematopoietic stem cells by decreasing Kit-Ras/cAMP-CREB pathway networks. FASEB J 33(5):5903", 2019,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC6463920/", "READ", "Mouse bone-marrow Lin- HSPC; actual spaceflight + RWV", "Space",
     "Reduced HSPC proliferation, G0/G1 arrest; differentiation not changed. No fold magnitude stated.", "contradictory", "r_g"),
    ("S13", "Proliferation of human hematopoietic bone marrow cells in simulated microgravity (PubMed 11332740)", "not extracted",
     "https://pubmed.ncbi.nlm.nih.gov/11332740/", "SNIPPET", "Human bone-marrow CD34+; simulated microgravity", "Space",
     "Normal-gravity cells grew up to ~3-fold by day 4-6; simulated-microgravity cells did not increase (ratio ~0.33).", "contradictory", "r_g"),
    ("S14", "Proliferation and functional characterization of human CD34+ bone marrow cells cultured in simulated microgravity. Experimental Hematology", "not extracted",
     "https://www.sciencedirect.com/science/article/abs/pii/S0301472X00003921", "SNIPPET", "Human CD34+; simulated microgravity (RWV)", "Space",
     "Day 5 total cells 11 vs 4.1-fold (ratio ~2.7); CD34+ 2.6 vs 1.0; but NOD/SCID repopulation 2.1 vs 14.8 (worse function, ratio ~0.14).", "mixed", "r_g, r_q"),
    ("S16", "NASA: Growing Stem Cells in Space to Improve Cancer and Disease Treatments (InSPA StemCellEX-H2)", 2026,
     "https://www.nasa.gov/missions/station/iss-research/growing-stem-cells-in-space-to-improve-cancer-and-disease-treatments/", "SNIPPET", "Human cord blood CD34+ expansion on ISS", "Space",
     "Target-system expansion experiment; no published yield results as of research date.", "contextual (pending)", "r_g"),
    ("S17", "Fisher V et al. Analysis of the Recovery of Cryopreserved and Thawed CD34+ and CD3+ Cells Collected for Hematopoietic Transplantation. Transfusion", 2014,
     "https://pmc.ncbi.nlm.nih.gov/articles/PMC3983171/", "READ", "Human mobilized CD34+, cryopreserved", "Space",
     "CD34+ recovery 97.5% (related, n=86) and 98.8% (unrelated, n=14); total-cell viability 64-79%.", "supportive (retention ceiling)", "s"),
    ("S18", "Post-thaw recovery of cryopreserved leukapheresis products (PubMed 37460360)", "not extracted",
     "https://pubmed.ncbi.nlm.nih.gov/37460360/", "SNIPPET", "Human apheresis CD34+, cryopreserved", "Space",
     "Median viable CD34+ recovery 73.7% autologous, 68.1% allogeneic.", "supportive (retention)", "s"),
    ("S19", "ACA Pharma: Omisirge WAC price (citing Gamida Cell order sheet)", 2026,
     "https://acapharma.net/omisirge-wac-price/", "SNIPPET", "Omidubicel list price (distributor page)", "Earth",
     "WAC $693,900 per treatment. Price, not cost; context only.", "contextual (price)", "none (context)"),
]

# ---------------------------------------------------------------------------
# Parameter tables. dist: 'fixed' or 'triangular'. Point estimate = mode. Low/High = range found in (or judged from) evidence.
# conf = author's SUBJECTIVE probability that the true long-run value lies inside [low, high]. NOT used by the model.
# ---------------------------------------------------------------------------
_target = None


def add(pid, label, unit, side, low, mode, high, dist, status, sources, basis, limits, conf, in_engine=True, formula=None):
    _target.append(dict(id=pid, label=label, unit=unit, side=side, low=low, mode=mode, high=high, dist=dist, status=status,
                        sources=sources, basis=basis, limits=limits, conf=conf, in_engine=in_engine, formula=formula))


# ======================= CM: iPSC-derived cardiomyocytes (team thesis) =======================
CM_P = []
_target = CM_P
add("U", "Cells per reporting unit", "cells/unit", "Both", 1e6, 1e6, 1e6, "fixed", "definition", "",
    "1 unit = 1 million correct, working heart muscle cells (cTnT+ cardiomyocytes passing release). A clinical dose is ~50-150 units (C11).", "Accounting choice only.", None)
add("N0", "Input iPSCs per differentiation batch", "cells", "Both", 1e9, 1e9, 1e9, "fixed", "measured in target system (derived)", "C12",
    "Industrial batch scale: ~1e9 input iPSCs gives ~2.5e9 CMs (C3 ratio); Ncardia reports 7.5e9-cell batches (C12).",
    "Assumes an orbital bioreactor can run the same batch size, which has NOT been demonstrated (favors space).", 0.5)
add("g_E", "Earth: total cells out per input iPSC (expansion during differentiation)", "fold", "Earth", 1.5, 2.8, 4, "triangular", "measured in target system (derived)", "C3, C2",
    "~2.5 CMs per input iPSC at ~90% purity -> ~2.8 total cells per input (C3). Range is judgment around that.", "Derived; per-protocol variation large.", 0.6)
add("q_E", "Earth: share of output that is correct heart muscle cells (purity)", "fraction", "Earth", 0.58, 0.88, 0.94, "triangular", "measured in target system", "C1, C2, C3, C12",
    "Bioreactor purity: 58% mean for one iPSC line (C1, low); ~87-90% typical (C3, C12); 93.5% best (C2, high).",
    "Earth's best 3D process already near 90%+, so there is little room for a purity gain.", 0.8)
add("p_E", "Earth: chance a batch succeeds and is released", "probability", "Earth", 0.75, 0.85, 0.95, "triangular", "measured in target system", "C4, C1",
    "'>85% purity in over 80% of runs' (C4); all 3 runs from batch-expanded cells failed in C1.", "No published GMP lot-release rate for iPSC-CMs.", 0.6)
add("E_cells", "Earth: starting iPSCs (master cell bank vial)", "USD/batch", "Earth", 1000, 5000, 10000, "triangular", "measured in target system", "C5",
    "GMP iPSC bank vial ~$1,000-10,000 (C5).", "Excludes $3M+ process development and bank creation costs.", 0.7)
add("E_consumables", "Earth: media, growth factors, single-use plastics", "USD/batch", "Earth", 10000, 40000, 120000, "triangular", "engineering estimate (derived)", "C3",
    "Research-grade materials ~$569 per 124M CMs (C3) -> ~$11k for ~2.5e9 CMs; GMP-grade media assumed 2-10x higher.", "GMP multiplier is judgment.", 0.5)
add("E_labor", "Earth: GMP labor", "USD/batch", "Earth", 50000, 120000, 250000, "triangular", "engineering estimate", "C5",
    "Labor is ~50% of hPSC manufacturing cost (C5); sized roughly equal to all other per-batch costs.", "Derived share, not a measured per-batch figure.", 0.5)
add("E_facility", "Earth: GMP clean-room allocation", "USD/batch", "Earth", 44000, 73000, 110000, "triangular", "engineering estimate (derived)", "C5",
    "Facility upkeep ~$2.2M/yr (C5) / 20-50 batches per year.", "Batches/year is judgment.", 0.55)
add("E_equipment", "Earth: bioreactor and equipment allocation", "USD/batch", "Earth", 5000, 15000, 40000, "triangular", "engineering estimate", "",
    "Stirred-tank bioreactor system amortized over batches. Author estimate.", "No source extracted.", 0.4)
add("E_qc", "Earth: release testing", "USD/batch", "Earth", 10000, 20000, 50000, "triangular", "measured in target system", "C5",
    "Outsourced testing ~$10k per lot (C5); more assays for a final cell product.", "", 0.6)
add("E_other", "Earth: other production cost", "USD/batch", "Earth", 0, 0, 0, "fixed", "not extracted", "", "All costs itemized above.", "", None)
add("E_delivery", "Earth: delivery cost", "USD/batch", "Earth", 0, 0, 0, "fixed", "not extracted", "",
    "Not modeled. Earth shipping losses also not modeled.", "Slightly favors Earth.", None)
add("r_g", "Space vs Earth: total cells out per input iPSC", "ratio", "Space", 0.8, 1.5, 4.0, "triangular", "measured in another system", "C6, C7, C9, C10",
    "1.5-4x more CMs per input cell under SIMULATED microgravity vs 3D at 1g (C6); much higher proliferation on ISS (C7); low end allows for slowed differentiation (C9, C10). Mode 1.5 = low end of C6 (judgment).",
    "Main evidence is simulated microgravity; real-flight yield per input cell not published.", 0.4)
add("r_q", "Space vs Earth: share of correct heart cells (purity ratio) - THESIS LEVER", "ratio", "Space", 0.9, 1.03, 1.06, "triangular", "measured in another system", "C6, C10, C13",
    "C6 reports ~99% purity in simulated microgravity; Earth best is ~88-94%, so the ratio can only be ~1.0-1.06 before purity hits 100%. Low end: differentiation inhibited (C10), radiation stress (C13).",
    "Cap: max(q_E) x max(r_q) must be <= 1. A large purity gain is only possible if Earth purity is low.", 0.5)
add("s", "Space: share of good cells surviving return, storage, thaw", "fraction", "Space", 0.5, 0.75, 0.97, "triangular", "measured in another system", "S17, S18, C8",
    "Post-thaw recovery of cryopreserved stem cells ~68-98% (S17, S18); low 0.5 allows for re-entry stress.", "CD34+ data used as proxy; Earth product cryo losses not modeled (favors Earth).", 0.5)
add("p_S", "Space: chance a batch succeeds and is released", "probability", "Space", 0.5, 0.8, 0.95, "triangular", "expert judgment", "C13",
    "No orbital GMP data. Includes launch delays, hardware faults, contamination, radiation-induced release failures.", "SPECULATIVE.", 0.35)
add("S_cells", "Space: starting iPSCs", "USD/batch", "Space", 1000, 5000, 10000, "triangular", "measured in target system", "C5", "Same as Earth.", "Sampled independently of Earth.", 0.7)
add("S_consumables", "Space: media and consumables flown", "USD/batch", "Space", 10000, 40000, 120000, "triangular", "engineering estimate (derived)", "C3", "Same as Earth.", "Sampled independently of Earth.", 0.5)
add("crew_hours", "Space: crew time per batch", "hours/batch", "Space", 0.5, 3, 10, "triangular", "engineering estimate", "",
    "Automated bioreactor; crew installs, swaps media modules, transfers to cold stowage. Author estimate.", "SPECULATIVE; full automation could approach zero.", 0.4, in_engine=False)
add("crew_rate", "Space: ISS crew time price", "USD/hour", "Space", 130000, 130000, 130000, "fixed", "measured in target system (price)", "S9",
    "NASA 2021 full-value commercial pricing.", "Policy price; NASA may discount research.", 0.8, in_engine=False)
add("S_labor", "Space: crew and operations time", "USD/batch", "Space", None, None, None, "triangular", "engineering estimate (derived)", "S9",
    "= crew hours x crew price.", "", None, formula=("crew_hours", "crew_rate"))
add("S_facility", "Space: orbital lab fee / implementation partner", "USD/batch", "Space", 50000, 150000, 500000, "triangular", "speculative scenario", "",
    "No public per-batch price (Space Tango, Redwire, BioServe quote on request). Author placeholder.", "SPECULATIVE.", 0.3)
add("S_equipment", "Space: flight bioreactor allocation", "USD/batch", "Space", 20000, 75000, 250000, "triangular", "speculative scenario", "",
    "Flight-qualified bioreactor amortized over batches. Author placeholder.", "SPECULATIVE.", 0.3)
add("S_qc", "Space: release testing", "USD/batch", "Space", 10000, 20000, 50000, "triangular", "measured in target system", "C5", "Same as Earth.", "", 0.6)
add("S_other", "Space: ground GMP work (pre-flight expansion, post-return processing)", "USD/batch", "Space", 60000, 150000, 300000, "triangular", "engineering estimate", "C5",
    "Ground team still expands iPSCs before flight and processes cells after return (labor + facility share).", "Judgment based on Earth labor/facility lines.", 0.45)
add("L", "Launch price to orbital lab", "USD/kg", "Space", 6500, 20000, 71800, "triangular", "measured in target system (price)", "S9, S10, S11",
    "Low: rideshare $6,500/kg (free-flyer orbit); mode: NASA ISS upmass $20,000/kg; high: CRS-2 ~$71,800/kg.", "Prices, not costs.", 0.75)
add("M", "Launched mass per batch (hardware share, media, packaging)", "kg/batch", "Space", 10, 25, 60, "triangular", "engineering estimate", "",
    "~1e9 iPSCs need ~1 L culture plus several litres of differentiation media over ~2 weeks, plus hardware share. Author estimate.", "SPECULATIVE.", 0.4)
add("ret_mass", "Return mass per batch incl. cold stowage", "kg/batch", "Space", 2, 5, 15, "triangular", "engineering estimate", "",
    "Cryo/cold-stowage container for one batch. Author estimate.", "SPECULATIVE.", 0.5, in_engine=False)
add("down_rate", "Downmass price", "USD/kg", "Space", 40000, 40000, 40000, "fixed", "measured in target system (price)", "S9", "NASA 2021 full-value downmass price.", "Policy price.", 0.8, in_engine=False)
add("C_return", "Return and recovery cost", "USD/batch", "Space", None, None, None, "triangular", "engineering estimate (derived)", "S9",
    "= return mass x downmass price.", "", None, formula=("ret_mass", "down_rate"))
add("C_other_del", "Other space delivery cost", "USD/batch", "Space", 0, 0, 0, "fixed", "not extracted", "", "Not modeled separately.", "", None)

CM_CLAIMS = [
    ("A1", "Real microgravity yields more correct heart cells per input iPSC than the best Earth 3D bioreactor (r_g x r_q > 1)", 0.35,
     "C6 (1.5-4x, simulated), C7 (proliferation on ISS)", "C9, C10 (differentiation suppressed)", "Real-flight yield per input cell not published."),
    ("A2", "Purity gain over the best Earth process exceeds 5% (r_q > 1.05)", 0.15,
     "C6 (~99% purity, simulated)", "C2, C3 (Earth already ~90%+)", "Math cap: purity cannot exceed 100%."),
    ("A3", "Microgravity does not impair the directed differentiation step", 0.45, "C6", "C9, C10", "Thesis needs pluripotency kept AND differentiation to work."),
    ("A4", "An orbital batch succeeds and is released at least 80% of the time", 0.35, "", "C13 (radiation/mutation risk)", "No orbital GMP data."),
    ("A5", "An orbital bioreactor can run an industrial batch (~1e9 input cells) within 5 years", 0.20, "", "", "Current ISS hardware is far smaller."),
    ("A6", "Launch to an orbital lab costs <= $20,000/kg for this payload", 0.60, "S9, S10", "S11", "Depends on provider and NASA policy."),
    ("A7", "Radiation over a 2-3 week run does not add release-blocking genetic damage", 0.60, "", "C13 (~5x mutations in HSPC on ISS)", "Release testing would catch some."),
]

# ======================= CMV: space viability case (same biology, future logistics) =======================
# Starts from CM and changes ONLY the rows below. Every changed row is marked 'projection' (or sits inside today's
# evidence range) so the calculator flags it as a bet rather than a measured fact.
import copy  # noqa: E402

CMV_OVERRIDES = {
    "r_g": dict(low=1.5, mode=3.0, high=4.0, status="measured in another system", sources="C6, C7",
                basis="Upper half of the simulated-microgravity evidence: 1.5-4x more heart cells per input stem cell vs 3D at 1g (C6); strong proliferation on ISS (C7). Mode 3.0 = midpoint of C6's range.",
                limits="BET: real-flight yield per input cell not yet published. Today's broader range is 0.8-4 (CM scenario).", conf=0.35),
    "r_q": dict(low=1.0, mode=1.05, high=1.06, sources="C6",
                basis="Near the 100%-purity ceiling: ~99% purity reported in simulated microgravity (C6).", limits="Small effect by construction: Earth purity already ~88%.", conf=0.4),
    "p_S": dict(low=0.7, mode=0.85, high=0.95, status="projection", sources="",
                basis="Mature, repeated autonomous flights (assumed).", limits="BET: no orbital GMP batch history.", conf=0.3),
    "crew_hours": dict(low=0, mode=0, high=0, status="projection", sources="P2",
                       basis="Autonomous free-flyer (Varda-style): no astronaut time.", limits="", conf=0.6),
    "S_labor": dict(low=10000, mode=30000, high=80000, status="projection", sources="P2", formula=None,
                    basis="Ground operations team monitoring an autonomous free-flyer instead of ISS crew time.", limits="BET: no published operations cost.", conf=0.4),
    "S_facility": dict(low=30000, mode=80000, high=200000, status="projection", sources="P2",
                       basis="Per-batch share of a free-flyer platform (bus, power, thermal) flown many times.", limits="BET: no published price.", conf=0.3),
    "S_equipment": dict(low=10000, mode=40000, high=120000, status="projection", sources="",
                        basis="Reusable flight bioreactor amortized over many flights.", limits="BET: placeholder.", conf=0.3),
    "L": dict(low=200, mode=1000, high=6500, status="projection", sources="P1, S10",
              basis="Low: Starship target ~$100-200/kg (aspirational, P1); mode: $1,000/kg = top of 'realistic mid-decade' range (P1); high: today's rideshare $6,500/kg (S10).",
              limits="BET: Starship has no published commercial price.", conf=0.45),
    "ret_mass": dict(low=5, mode=5, high=5, dist="fixed", conf=None, basis="Not used: return priced per capsule share below.", limits=""),
    "C_return": dict(low=50000, mode=100000, high=250000, status="projection", sources="P2", formula=None,
                     basis="~$5M per reentry capsule flight (P2) shared by ~20-100 batches (50 at the mode).",
                     limits="BET: batches per capsule assumed; capsule capacity not published.", conf=0.35),
}
CMV_P = copy.deepcopy(CM_P)
for _p in CMV_P:
    if _p["id"] in CMV_OVERRIDES:
        _p.update(CMV_OVERRIDES[_p["id"]])
        if _p["low"] == _p["high"]:
            _p["dist"] = "fixed"

CMV_CLAIMS = [
    ("B1", "Real microgravity gives ~3x more correct heart cells per input stem cell than the best Earth 3D bioreactor", 0.25,
     "C6 (1.5-4x, simulated microgravity), C7 (ISS proliferation)", "C9, C10 (differentiation suppressed)", "The single biggest bet. Test it on the next flight."),
    ("B2", "Launch to LEO falls to ~$1,000/kg or less for small biotech payloads by the early 2030s", 0.50, "P1", "", "Starship price unpublished."),
    ("B3", "An autonomous free-flyer can run a GMP cell-differentiation batch without crew", 0.35, "P2 (autonomous capsules fly today)", "", "No cell-therapy batch has flown this way."),
    ("B4", "A reentry capsule flight (~$5M) can be shared by ~50 batches", 0.30, "P2", "", "Requires high demand and capsule capacity."),
    ("B5", "Space-made cells pass the same release tests after re-entry (same quality as Earth cells)", 0.45, "C8 (CM structure unchanged)", "C13 (radiation)", "Equivalence must be shown, not assumed."),
]

# ======================= HSC: cord-blood CD34+ expansion =======================
HSC_P = []
_target = HSC_P
add("U", "Cells per reporting unit", "cells/unit", "Both", 1e6, 1e6, 1e6, "fixed", "definition", "",
    "1 unit = 1 million functional HSC-equivalent (CD34+) cells. Not a clinical dose.", "Accounting choice only.", None)
add("N0", "Viable starting CD34+ cells per batch (one cord blood unit)", "cells", "Both", 5e6, 5e6, 5e6, "fixed", "measured in target system (derived)", "S2",
    "Derived: ~6.6e8 CD34+ per omidubicel product / 130-fold median expansion = ~5e6 starting CD34+.", "Derived arithmetic from snippet values; cord units vary widely.", 0.6)
add("g_E", "Earth gross CD34+ expansion factor", "fold", "Earth", 35, 130, 330, "triangular", "measured in target system", "S1, S2, S3",
    "UM171 ~35-fold (low), omidubicel median 130-fold (mode), SR-1 median 330-fold (high).", "Different protocols/products; medians, not long-run means; triangle shape is judgment.", 0.85)
add("q_E", "Earth fraction of expanded CD34+ meeting functional qualification", "fraction", "Earth", 0.5, 0.7, 0.8, "triangular", "speculative scenario", "",
    "No post-expansion functional-fraction value was extracted. Placeholder range by author judgment.", "NOT EVIDENCE-BASED. Needs an assay definition.", 0.4)
add("p_E", "Earth batch release probability", "probability", "Earth", 0.88, 0.96, 0.99, "triangular", "measured in target system", "S1, S4, S8",
    "SR-1: 2/20 failed release (90%); UM171: 26/27 expanded (96%); CAR-T failure ~4% (other system).", "Small samples; failure rates fell as processes matured.", 0.8)
add("E_cells", "Earth: starting material (cord blood unit)", "USD/batch", "Earth", 30000, 36000, 60000, "triangular", "measured in another system", "S7",
    "Cord blood unit procurement $30k-$60k+; mode near reported ~$36k mean (vendor snippet).", "Price years vary; no inflation adjustment.", 0.75)
for _id, _lab in (("E_consumables", "media, cytokines, consumables"), ("E_labor", "labor"), ("E_facility", "GMP facility allocation"),
                  ("E_equipment", "equipment allocation"), ("E_qc", "quality testing / release")):
    add(_id, "Earth: " + _lab, "USD/batch", "Earth", 0, 0, 0, "fixed", "not extracted (included in COGS line)", "S6",
        "Included in the manufacturing COGS line to avoid double counting; breakdown not extracted.", "Breakdown unavailable.", None)
add("E_other", "Earth: expansion manufacturing COGS excl. starting material (labor, facility, QC, consumables)", "USD/batch", "Earth", 40000, 100000, 220000, "triangular", "measured in another system", "S5, S6",
    "Autologous CAR-T COGS ~$96k/dose (S5); reviews cite $170k-220k per batch, projected to fall to $40k-70k.", "CAR-T, not HSC expansion; 2019 USD.", 0.6)
add("E_delivery", "Earth: delivery cost", "USD/batch", "Earth", 0, 0, 0, "fixed", "not extracted", "",
    "Not modeled (courier cost small relative to COGS). Earth shipping losses also not modeled.", "Slightly favors Earth.", None)
add("r_g", "Space/Earth gross expansion ratio", "ratio", "Space", 0.3, 1.0, 2.7, "triangular", "measured in another system", "S12, S13, S14, S16",
    "Low ~0.33 (sim-microgravity human BM CD34+, no growth vs 3-fold); high ~2.7 (sim-microgravity total cells 11 vs 4.1-fold); mode 1.0 (early ISS HSC work reported comparable growth). Mouse spaceflight shows decreases.",
    "No human cord-blood CD34+ expansion measured in real microgravity under manufacturing conditions.", 0.5)
add("r_q", "Space/Earth qualification-fraction ratio", "ratio", "Space", 0.15, 1.0, 1.25, "triangular", "measured in another system", "S14, C13",
    "Low ~0.14 (NOD/SCID repopulation 2.1 vs 14.8 after simulated microgravity); ISS HSPC aging signatures (C13); high 1.25 speculative.",
    "Upper bound speculative. max(q_E) x max(r_q) = 1.0 keeps q_S <= 1.", 0.5)
add("s", "Post-return retention of qualified space output", "fraction", "Space", 0.5, 0.75, 0.97, "triangular", "measured in another system", "S17, S18",
    "Post-thaw CD34+ recovery: ~98% (S17) down to ~68-74% (S18); low 0.5 allows for re-entry stress.", "Earth product cryo/shipping losses not modeled, which favors Earth.", 0.6)
add("p_S", "Space batch release probability", "probability", "Space", 0.6, 0.85, 0.96, "triangular", "expert judgment", "",
    "No orbital GMP batch data. Includes launch scrubs, hardware faults, contamination, return delays. Author judgment.", "SPECULATIVE.", 0.4)
add("S_cells", "Space: starting material (cord blood unit)", "USD/batch", "Space", 30000, 36000, 60000, "triangular", "measured in another system", "S7", "Same as Earth.", "Sampled independently of Earth.", 0.75)
add("S_consumables", "Space: media, cytokines, consumables", "USD/batch", "Space", 0, 0, 0, "fixed", "not extracted (included in ground COGS line)", "", "Included in ground COGS line.", "", None)
add("crew_hours", "Space: crew time per batch", "hours/batch", "Space", 0.5, 2, 8, "triangular", "engineering estimate", "",
    "Automated bioreactor payload; crew handles install, sampling and cold-stow transfer. Author estimate.", "SPECULATIVE.", 0.5, in_engine=False)
add("crew_rate", "Space: ISS crew time price", "USD/hour", "Space", 130000, 130000, 130000, "fixed", "measured in target system (price)", "S9",
    "NASA 2021 full-value commercial pricing.", "Policy price.", 0.8, in_engine=False)
add("S_labor", "Space: operations / crew time", "USD/batch", "Space", None, None, None, "triangular", "engineering estimate (derived)", "S9",
    "= crew hours x crew price.", "", None, formula=("crew_hours", "crew_rate"))
add("S_facility", "Space: orbital facility / implementation-partner allocation", "USD/batch", "Space", 50000, 150000, 500000, "triangular", "speculative scenario", "",
    "No public per-batch price. Author placeholder.", "SPECULATIVE.", 0.3)
add("S_equipment", "Space: flight hardware allocation", "USD/batch", "Space", 20000, 75000, 250000, "triangular", "speculative scenario", "",
    "Flight-qualified bioreactor amortized over expected batches. Author placeholder.", "SPECULATIVE.", 0.3)
add("S_qc", "Space: quality testing / release", "USD/batch", "Space", 0, 0, 0, "fixed", "not extracted (included in ground COGS line)", "", "Included in ground COGS line.", "", None)
add("S_other", "Space: ground GMP processing COGS (pre-flight prep, post-return QC)", "USD/batch", "Space", 40000, 100000, 220000, "triangular", "measured in another system", "S5, S6",
    "Assumed equal to Earth COGS range.", "Sampled independently of Earth COGS.", 0.5)
add("L", "Launch price to LEO station", "USD/kg", "Space", 6500, 20000, 71800, "triangular", "measured in target system (price)", "S9, S10, S11",
    "Low: rideshare $6,500/kg; mode: NASA ISS upmass $20,000/kg; high: CRS-2 ~$71,800/kg.", "Prices, not costs.", 0.75)
add("M", "Allocated launched mass per batch", "kg/batch", "Space", 5, 15, 40, "triangular", "engineering estimate", "",
    "CubeLab-scale bioreactor share + media + cold-chain packaging. Author estimate.", "SPECULATIVE.", 0.5)
add("ret_mass", "Return (downmass) per batch incl. cold-stowage", "kg/batch", "Space", 2, 5, 15, "triangular", "engineering estimate", "",
    "Cryo/cold-stowage container for one batch. Author estimate.", "SPECULATIVE.", 0.5, in_engine=False)
add("down_rate", "Downmass price", "USD/kg", "Space", 40000, 40000, 40000, "fixed", "measured in target system (price)", "S9", "NASA 2021 full-value downmass price.", "Policy price.", 0.8, in_engine=False)
add("C_return", "Return and recovery cost", "USD/batch", "Space", None, None, None, "triangular", "engineering estimate (derived)", "S9",
    "= return mass x downmass price.", "", None, formula=("ret_mass", "down_rate"))
add("C_other_del", "Other space delivery cost", "USD/batch", "Space", 0, 0, 0, "fixed", "not extracted", "", "Not modeled separately.", "", None)

HSC_CLAIMS = [
    ("A1", "Real microgravity increases gross expansion of human cord-blood CD34+ cells vs matched Earth controls (r_g > 1)", 0.30,
     "S14 (sim. microgravity, more cells at day 5)", "S12, S13 (decreased proliferation), C13", "Mixed evidence; mostly simulated microgravity or mouse cells."),
    ("A2", "Microgravity-expanded cells keep Earth-equivalent function (r_q >= 1)", 0.25,
     "None direct", "S14 (repopulation 2.1 vs 14.8), C13 (ISS aging signatures)", "Functional data point the other way."),
    ("A3", "Earth CD34+ expansion of ~35-330 fold is achievable at GMP scale", 0.85, "S1, S2, S3", "", "Approved product (Omisirge) exists."),
    ("A4", "Orbital batch release probability >= 85%", 0.35, "", "No orbital GMP data", "Pure judgment."),
    ("A5", "Launch to an orbital lab at <= $20,000/kg is available for this payload class", 0.60, "S9, S10", "S11", "Depends on provider and NASA policy."),
    ("A6", "Post-return retention of qualified cells >= 75%", 0.50, "S17", "S18", "Cryo recovery data from Earth shipping, not re-entry."),
    ("A7", "Space-specific overheads (facility, hardware, crew) <= $300k per batch at the point estimate", 0.40, "", "", "No public prices; placeholders."),
]


def resolve(P):
    by = {p["id"]: p for p in P}
    for p in P:  # derived products (also written as formulas in the workbook)
        if p["formula"]:
            a, b = (by[k] for k in p["formula"])
            p["low"], p["mode"], p["high"] = a["low"] * b["low"], a["mode"] * b["mode"], a["high"] * b["high"]
    return by


SCEN = {
    "CMV": dict(P=CMV_P, by=resolve(CMV_P), claims=CMV_CLAIMS, js="SCViable",
                title="Space viability case: heart muscle cells on an autonomous free-flyer (projection)",
                unit="unit (1 million correct heart muscle cells)"),
    "CM": dict(P=CM_P, by=resolve(CM_P), claims=CM_CLAIMS, js="SCiPSC",
               title="Team thesis: iPSC-derived heart muscle cells (cardiomyocytes)",
               unit="unit (1 million correct heart muscle cells)"),
    "HSC": dict(P=HSC_P, by=resolve(HSC_P), claims=HSC_CLAIMS, js="SCLiterature",
                title="Blood stem cells: cord-blood CD34+ expansion",
                unit="unit (1 million qualified CD34+ cells)"),
}
# Backward-compatible aliases (HSC was the original scenario)
P, BY_ID, CLAIMS = HSC_P, SCEN["HSC"]["by"], HSC_CLAIMS


def model(v):
    """Independent Python implementation of PRD section 6/7 equations (mirrors the spreadsheet formulas)."""
    g_S, q_S = v["g_E"] * v["r_g"], v["q_E"] * v["r_q"]
    Y_E = v["N0"] * v["g_E"] * v["q_E"] / v["U"]
    Y_S_pre = v["N0"] * g_S * q_S / v["U"]
    Q_E, Q_S_pre, Q_S = v["p_E"] * Y_E, v["p_S"] * Y_S_pre, v["p_S"] * Y_S_pre * v["s"]
    C_E_prod = sum(v[k] for k in ("E_cells", "E_consumables", "E_labor", "E_facility", "E_equipment", "E_qc", "E_other"))
    C_S_prod = sum(v[k] for k in ("S_cells", "S_consumables", "S_labor", "S_facility", "S_equipment", "S_qc", "S_other"))
    C_E_del = C_E_prod + v["E_delivery"]
    C_launch = v["L"] * v["M"]
    C_S_del = C_S_prod + C_launch + v["C_return"] + v["C_other_del"]
    inf = float("inf")
    c_E = C_E_del / Q_E if Q_E > 0 else inf
    c_S = C_S_del / Q_S if Q_S > 0 else inf
    out = dict(Q_E=Q_E, Q_S_pre=Q_S_pre, Q_S=Q_S, C_E_prod=C_E_prod, C_S_prod=C_S_prod, C_E_del=C_E_del, C_S_del=C_S_del,
               c_E=c_E, c_S=c_S, c_E_prod=C_E_prod / Q_E if Q_E > 0 else inf, c_S_prod=C_S_prod / Q_S_pre if Q_S_pre > 0 else inf)
    out["savings"] = 1 - c_S / c_E if (Q_E > 0 and Q_S > 0 and c_E > 0) else None
    out["L_max"] = (c_E * Q_S - C_S_prod - v["C_return"] - v["C_other_del"]) / v["M"] if v["M"] > 0 else None
    den = c_E * v["p_S"] * v["N0"] * v["g_E"] * q_S * v["s"]
    out["r_g_min"] = C_S_del * v["U"] / den if den > 0 else None
    denq = c_E * v["p_S"] * v["N0"] * g_S * v["q_E"] * v["s"]
    out["r_q_min"] = C_S_del * v["U"] / denq if denq > 0 else None
    out["r_q_cap"] = 1 / v["q_E"] if v["q_E"] > 0 else None
    budget = c_E * Q_S
    out["C_return_max"] = budget - C_S_prod - C_launch - v["C_other_del"]
    out["C_S_prod_max"] = budget - C_launch - v["C_return"] - v["C_other_del"]
    return out


def point_values(Plist=None):
    return {p["id"]: p["mode"] for p in (Plist or P) if p["in_engine"]}


def tri_exceed(x, a, c, b):
    if x <= a:
        return 1.0
    if x >= b:
        return 0.0
    if x <= c:
        return 1 - (x - a) ** 2 / ((b - a) * (c - a))
    return (b - x) ** 2 / ((b - a) * (b - c))


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------
FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
BLACK = Font(name=FONT)
GREEN = Font(name=FONT, color="008000")
BOLD = Font(name=FONT, bold=True)
TITLE = Font(name=FONT, bold=True, size=14)
HEAD_FILL = PatternFill("solid", fgColor="DDE6F0")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
WRAP = Alignment(wrap_text=True, vertical="top")
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment = BOLD, HEAD_FILL, WRAP
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def put(ws, r, label, formula, fmt=None, unit="", note="", bold=False, link=False):
    ws.cell(row=r, column=1, value=label).font = BOLD if bold else BLACK
    ws.cell(row=r, column=1).alignment = WRAP
    c = ws.cell(row=r, column=2, value=formula)
    c.font = GREEN if link else (BOLD if bold else BLACK)
    if fmt:
        c.number_format = fmt
    ws.cell(row=r, column=3, value=unit).font = BLACK
    ws.cell(row=r, column=4, value=note).font = BLACK
    ws.cell(row=r, column=4).alignment = WRAP
    return f"B{r}"


def build_scenario(wb, pre, sc):
    P, by = sc["P"], sc["by"]
    A_ = f"{pre}_Assumptions"
    wa = wb.create_sheet(A_)
    wa["A1"], wa["A1"].font = f"{sc['title']}: assumptions. Point estimate = Most likely. Blue = input, black = formula.", TITLE
    wa["A2"], wa["A2"].font = "Low/High = range found in (or judged from) the evidence. Subjective confidence = author's probability the true value lies in [Low, High]; not used by the model.", BLACK
    cols = ["ID", "Parameter", "Unit", "Side", "Low", "Most likely (point)", "High", "Distribution", "Evidence status",
            "Evidence strength", "Source IDs", "Basis / derivation", "Limitations", "Subjective confidence range holds", "Used directly by calculator?"]
    header(wa, 4, cols, [12, 46, 13, 8, 13, 15, 13, 12, 26, 12, 14, 60, 45, 14, 12])
    strength = {"measured in target system": "Strong", "measured in target system (derived)": "Strong", "measured in target system (price)": "Strong",
                "measured in another system": "Moderate", "engineering estimate": "Weak", "engineering estimate (derived)": "Weak",
                "expert judgment": "Very weak", "speculative scenario": "Very weak", "projection": "Projection (future bet)"}
    row_of = {p["id"]: 5 + i for i, p in enumerate(P)}
    for p in P:
        r = row_of[p["id"]]
        vals = [p["id"], p["label"], p["unit"], p["side"]]
        for col, val in enumerate(vals, 1):
            wa.cell(row=r, column=col, value=val).font = BLACK
        for col, key in ((5, "low"), (6, "mode"), (7, "high")):
            c = wa.cell(row=r, column=col)
            if p["formula"]:
                a, b = p["formula"]
                L = get_column_letter(col)
                c.value, c.font = f"={L}{row_of[a]}*{L}{row_of[b]}", BLACK
            else:
                c.value, c.font = p[key], BLUE
            if p["unit"].startswith("USD"):
                c.number_format = USD
            elif p["unit"] in ("cells", "cells/unit"):
                c.number_format = "0.00E+00"
            elif p["unit"] in ("fraction", "probability"):
                c.number_format = "0.0%"
        rest = [p["dist"], p["status"], strength.get(p["status"], "n/a"), p["sources"], p["basis"], p["limits"]]
        for col, val in enumerate(rest, 8):
            wa.cell(row=r, column=col, value=val).font = BLACK
        cc = wa.cell(row=r, column=14, value=p["conf"])
        cc.font, cc.number_format = BLUE, "0%"
        wa.cell(row=r, column=15, value="yes" if p["in_engine"] else "via derived row").font = BLACK
        for col in range(1, 16):
            wa.cell(row=r, column=col).alignment = WRAP
        if p["id"] in ("r_g", "r_q", "q_E", "L", "s", "p_S"):
            wa.cell(row=r, column=6).fill = KEY_FILL
    wa.freeze_panes = "C5"
    wa.cell(row=row_of["r_q"], column=6).comment = Comment("Thesis lever: purity in space vs Earth. Cannot push space purity above 100%.", "model")

    def A(pid):
        return f"{A_}!$F${row_of[pid]}"

    # Earth sheet
    E_ = f"{pre}_Earth"
    we = wb.create_sheet(E_)
    we["A1"], we["A1"].font = f"{sc['title']} - Earth: cost and output per attempted batch (point estimates)", TITLE
    header(we, 3, ["Line item", "Value", "Unit", "Note"], [62, 16, 16, 70])
    r, first = 4, 4
    for pid in ("E_cells", "E_consumables", "E_labor", "E_facility", "E_equipment", "E_qc", "E_other"):
        put(we, r, by[pid]["label"], f"={A(pid)}", USD, "USD/batch", by[pid]["status"], link=True)
        r += 1
    E = {}
    E["Cprod"] = put(we, r, "C_E_prod: production cost per batch", f"=SUM(B{first}:B{r-1})", USD, "USD/batch", bold=True); r += 1
    E["deliv"] = put(we, r, "Earth delivery cost", f"={A('E_delivery')}", USD, "USD/batch", link=True); r += 1
    E["Cdel"] = put(we, r, "C_E_del: delivered cost per batch", f"={E['Cprod']}+{E['deliv']}", USD, "USD/batch", bold=True); r += 2
    E["N0"] = put(we, r, "N0: starting cells", f"={A('N0')}", "0.00E+00", "cells", link=True); r += 1
    E["g"] = put(we, r, "g_E: growth (cells out per starting cell)", f"={A('g_E')}", "0.0", "fold", link=True); r += 1
    E["q"] = put(we, r, "q_E: share of correct / qualified cells", f"={A('q_E')}", "0.0%", "fraction", link=True); r += 1
    E["p"] = put(we, r, "p_E: batch success probability", f"={A('p_E')}", "0.0%", "probability", link=True); r += 1
    E["U"] = put(we, r, "U: cells per unit", f"={A('U')}", "0.00E+00", "cells/unit", link=True); r += 1
    E["Y"] = put(we, r, "Y_E = N0 x g_E x q_E / U  (successful batch)", f"={E['N0']}*{E['g']}*{E['q']}/{E['U']}", "#,##0.0", "units/batch"); r += 1
    E["Q"] = put(we, r, "Q_E = p_E x Y_E  (expected per attempted batch)", f"={E['p']}*{E['Y']}", "#,##0.0", "units/batch", bold=True); r += 1
    E["cprod"] = put(we, r, "c_E_prod: production cost per unit", f'=IF({E["Q"]}>0,{E["Cprod"]}/{E["Q"]},"No usable output")', USD2, "USD/unit"); r += 1
    E["c"] = put(we, r, "c_E: delivered cost per unit", f'=IF({E["Q"]}>0,{E["Cdel"]}/{E["Q"]},"No usable output")', USD2, "USD/unit", bold=True)
    we[E["c"]].fill = KEY_FILL

    # Space sheet
    S_ = f"{pre}_Space"
    wsp = wb.create_sheet(S_)
    wsp["A1"], wsp["A1"].font = f"{sc['title']} - Space (LEO): cost and output per attempted batch (point estimates)", TITLE
    header(wsp, 3, ["Line item", "Value", "Unit", "Note"], [62, 16, 16, 70])
    r, first = 4, 4
    for pid in ("S_cells", "S_consumables", "S_labor", "S_facility", "S_equipment", "S_qc", "S_other"):
        put(wsp, r, by[pid]["label"], f"={A(pid)}", USD, "USD/batch", by[pid]["status"], link=True)
        r += 1
    S = {}
    S["Cprod"] = put(wsp, r, "C_S_prod: production (orbital manufacturing) cost per batch", f"=SUM(B{first}:B{r-1})", USD, "USD/batch", bold=True); r += 1
    S["L"] = put(wsp, r, "L: launch price", f"={A('L')}", USD, "USD/kg", link=True); r += 1
    S["M"] = put(wsp, r, "M: launched mass", f"={A('M')}", "0.0", "kg/batch", link=True); r += 1
    S["launch"] = put(wsp, r, "C_launch = L x M", f"={S['L']}*{S['M']}", USD, "USD/batch"); r += 1
    S["ret"] = put(wsp, r, "Return and recovery (return mass x downmass price)", f"={A('C_return')}", USD, "USD/batch", link=True); r += 1
    S["oth"] = put(wsp, r, "Other delivery", f"={A('C_other_del')}", USD, "USD/batch", link=True); r += 1
    S["Cdel"] = put(wsp, r, "C_S_del: delivered cost per batch", f"={S['Cprod']}+{S['launch']}+{S['ret']}+{S['oth']}", USD, "USD/batch", bold=True); r += 2
    S["N0"] = put(wsp, r, "N0: starting cells (paired with Earth)", f"={A('N0')}", "0.00E+00", "cells", link=True); r += 1
    S["gE"] = put(wsp, r, "g_E: Earth growth", f"={A('g_E')}", "0.0", "fold", link=True); r += 1
    S["rg"] = put(wsp, r, "r_g: space/Earth growth ratio", f"={A('r_g')}", "0.00", "ratio", link=True); r += 1
    S["g"] = put(wsp, r, "g_S = g_E x r_g", f"={S['gE']}*{S['rg']}", "0.0", "fold"); r += 1
    S["qE"] = put(wsp, r, "q_E: Earth share of correct cells", f"={A('q_E')}", "0.0%", "fraction", link=True); r += 1
    S["rq"] = put(wsp, r, "r_q: space/Earth purity ratio (thesis lever)", f"={A('r_q')}", "0.00", "ratio", link=True); r += 1
    S["q"] = put(wsp, r, "q_S = q_E x r_q", f"={S['qE']}*{S['rq']}", "0.0%", "fraction"); r += 1
    S["s"] = put(wsp, r, "s: post-return retention", f"={A('s')}", "0.0%", "fraction", link=True); r += 1
    S["p"] = put(wsp, r, "p_S: batch success probability", f"={A('p_S')}", "0.0%", "probability", link=True); r += 1
    S["U"] = put(wsp, r, "U: cells per unit", f"={A('U')}", "0.00E+00", "cells/unit", link=True); r += 1
    S["Ypre"] = put(wsp, r, "Y_S_pre = N0 x g_S x q_S / U", f"={S['N0']}*{S['g']}*{S['q']}/{S['U']}", "#,##0.0", "units/batch"); r += 1
    S["Y"] = put(wsp, r, "Y_S = Y_S_pre x s", f"={S['Ypre']}*{S['s']}", "#,##0.0", "units/batch"); r += 1
    S["Qpre"] = put(wsp, r, "Q_S_pre = p_S x Y_S_pre", f"={S['p']}*{S['Ypre']}", "#,##0.0", "units/batch"); r += 1
    S["Q"] = put(wsp, r, "Q_S = p_S x Y_S  (expected delivered per attempted batch)", f"={S['p']}*{S['Y']}", "#,##0.0", "units/batch", bold=True); r += 1
    S["cprod"] = put(wsp, r, "c_S_prod: production cost per unit (before transport/return losses)", f'=IF({S["Qpre"]}>0,{S["Cprod"]}/{S["Qpre"]},"No usable output")', USD2, "USD/unit"); r += 1
    S["c"] = put(wsp, r, "c_S: delivered cost per unit", f'=IF({S["Q"]}>0,{S["Cdel"]}/{S["Q"]},"No usable output")', USD2, "USD/unit", bold=True)
    wsp[S["c"]].fill = KEY_FILL

    # Comparison sheet
    wc = wb.create_sheet(f"{pre}_Comparison")
    wc["A1"], wc["A1"].font = f"{sc['title']} - comparison and viability (deterministic, point estimates)", TITLE
    header(wc, 3, ["Metric", "Value", "Unit", "Note"], [74, 18, 14, 80])
    e = lambda k: f"{E_}!{E[k]}"
    s = lambda k: f"{S_}!{S[k]}"
    r = 4
    cE = put(wc, r, "Earth delivered cost per unit (c_E)", f"={e('c')}", USD2, "USD/unit", link=True); r += 1
    cS = put(wc, r, "Space delivered cost per unit (c_S)", f"={s('c')}", USD2, "USD/unit", link=True); r += 1
    put(wc, r, "Space production cost per unit (c_S_prod)", f"={s('cprod')}", USD2, "USD/unit", "Includes orbital facility costs; NOT proof of a microgravity effect.", link=True); r += 1
    put(wc, r, "Relative delivered savings = 1 - c_S/c_E", f'=IF(AND(ISNUMBER({cE}),ISNUMBER({cS}),{cE}>0),1-{cS}/{cE},"Not defined")', "0.0%", "", "Negative = space more expensive."); r += 1
    put(wc, r, "Verdict", f'=IF(AND(ISNUMBER({cE}),ISNUMBER({cS})),IF({cS}<{cE},"Space cheaper","Earth cheaper or equal"),"See no-output flags")', None, "", bold=True); r += 2
    wc.cell(row=r, column=1, value="AT WHAT COSTS IS IT VIABLE? (each line changes one input, others held at point estimates)").font = BOLD; r += 1
    budget = put(wc, r, "Max total space spend per batch to match Earth = c_E x Q_S", f'=IFERROR({cE}*{s("Q")},"Not defined")', USD, "USD/batch"); r += 1
    put(wc, r, "Max launch price L_max = (c_E x Q_S - C_S_prod - C_return - C_other) / M",
        f'=IF({s("M")}>0,({cE}*{s("Q")}-{s("Cprod")}-{s("ret")}-{s("oth")})/{s("M")},"Not applicable (M=0)")', USD, "USD/kg",
        "Negative = not competitive even with free launch."); r += 1
    put(wc, r, "Max return cost per batch", f'=IFERROR({budget}-{s("Cprod")}-{s("launch")}-{s("oth")},"Not defined")', USD, "USD/batch", "Negative = not viable even with free return."); r += 1
    put(wc, r, "Max orbital manufacturing cost per batch (C_S_prod)", f'=IFERROR({budget}-{s("launch")}-{s("ret")}-{s("oth")},"Not defined")', USD, "USD/batch", "Negative = not viable even with free manufacturing."); r += 1
    rgmin = put(wc, r, "Break-even growth ratio r_g_min = C_S_del x U / (c_E x p_S x N0 x g_E x q_S x s)",
                f'=IFERROR({s("Cdel")}*{s("U")}/({cE}*{s("p")}*{s("N0")}*{s("gE")}*{s("q")}*{s("s")}),"Not defined")', "0.00", "ratio"); r += 1
    rqmin = put(wc, r, "Break-even purity ratio r_q_min = C_S_del x U / (c_E x p_S x N0 x g_S x q_E x s)",
                f'=IFERROR({s("Cdel")}*{s("U")}/({cE}*{s("p")}*{s("N0")}*{s("g")}*{s("qE")}*{s("s")}),"Not defined")', "0.00", "ratio", "Thesis lever."); r += 1
    rqcap = put(wc, r, "Maximum possible purity ratio = 1 / q_E (space purity capped at 100%)", f'=IFERROR(1/{s("qE")},"Not defined")', "0.00", "ratio"); r += 1
    put(wc, r, "Can the purity advantage alone make space viable?", f'=IF(AND(ISNUMBER({rqmin}),ISNUMBER({rqcap})),IF({rqmin}<={rqcap},"Possible in principle","NO: required ratio exceeds the 100%-purity cap"),"Not defined")', None, "", bold=True); r += 1
    put(wc, r, "Space cost per unit if 100% of space cells were correct", f'=IFERROR({s("Cdel")}/({s("p")}*{s("N0")}*{s("g")}*{s("s")}/{s("U")}),"Not defined")', USD2, "USD/unit", "Best case for the thesis."); r += 1
    put(wc, r, "Required output ratio C_S_del / C_E_del (Q_S/Q_E must exceed this)", f'=IFERROR({s("Cdel")}/{e("Cdel")},"Not defined")', "0.00", "ratio"); r += 1
    put(wc, r, "Current output ratio Q_S / Q_E", f'=IFERROR({s("Q")}/{e("Q")},"Not defined")', "0.00", "ratio"); r += 2
    a, c_, b = (f"{A_}!${col}${row_of['r_g']}" for col in ("E", "F", "G"))
    x = rgmin
    p_rg = put(wc, r, "P(r_g > r_g_min) under the triangular r_g range, all else fixed",
               f'=IF(ISNUMBER({x}),IF({x}<={a},1,IF({x}>={b},0,IF({x}<={c_},1-({x}-{a})^2/(({b}-{a})*({c_}-{a})),({b}-{x})^2/(({b}-{a})*({b}-{c_}))))),"Not defined")',
               "0.00%", "probability", "Analytic triangular tail. Conditional on the range; not a forecast. Joint uncertainty: run the calculator's Monte Carlo.")

    # Probabilities sheet
    wp = wb.create_sheet(f"{pre}_Probabilities")
    wp["A1"], wp["A1"].font = f"{sc['title']} - key assumptions and the author's subjective probability each is true", TITLE
    wp["A2"] = ("Subjective judgments drafted from the evidence inventory for discussion; review and replace with your own. "
                "NOT derived from p-values, NOT used in the simulation, NOT independent of one another.")
    wp["A2"].font = BLACK
    header(wp, 4, ["ID", "Claim", "Subjective P(true)", "Evidence for", "Evidence against", "Comment"], [6, 70, 14, 30, 36, 50])
    r = 5
    for row in sc["claims"]:
        for col, val in enumerate(row, 1):
            c = wp.cell(row=r, column=col, value=val)
            c.font, c.alignment = (BLUE if col == 3 else BLACK), WRAP
        wp.cell(row=r, column=3).number_format = "0%"
        r += 1
    put(wp, r + 1, "Product of all P IF independent (they are not; illustration only)", f"=PRODUCT(C5:C{r-1})", "0.00%")
    put(wp, r + 2, f"Model-based P(r_g > break-even), from {pre}_Comparison", f"={pre}_Comparison!{p_rg}", "0.00%", link=True)
    return row_of


def build_prd_check(wb):
    wc = wb.create_sheet("PRD_Check")
    wc["A1"], wc["A1"].font = "PRD section 11 check case (hypothetical inputs; verifies the formulas)", TITLE
    header(wc, 3, ["Item", "Value", "Unit", "Note"], [60, 18, 10, 30])
    r = 4
    chk_inputs = [("U", 1e6), ("N0", 1e6), ("g_E", 20), ("r_g", 2), ("q_E", 0.5), ("r_q", 1), ("s", 0.8), ("p_E", 1), ("p_S", 1),
                  ("C_E_prod", 100000), ("E_delivery", 0), ("C_S_prod", 200000), ("L", 2000), ("M", 50), ("C_return", 50000), ("C_other", 0)]
    ref = {}
    for k, v in chk_inputs:
        wc.cell(row=r, column=1, value=k).font = BLACK
        c = wc.cell(row=r, column=2, value=v)
        c.font = BLUE
        ref[k] = f"B{r}"
        r += 1
    f = lambda k: ref[k]
    yE = put(wc, r, "Earth output Y_E (expect 10)", f"={f('N0')}*{f('g_E')}*{f('q_E')}/{f('U')}", "0.000"); r += 1
    ySpre = put(wc, r, "Space pre-return output (expect 20)", f"={f('N0')}*{f('g_E')}*{f('r_g')}*{f('q_E')}*{f('r_q')}/{f('U')}", "0.000"); r += 1
    yS = put(wc, r, "Space delivered output (expect 16)", f"={ySpre}*{f('s')}", "0.000"); r += 1
    chk_cE = put(wc, r, "Earth delivered cost (expect $10,000)", f"=({f('C_E_prod')}+{f('E_delivery')})/({f('p_E')}*{yE})", USD2); r += 1
    put(wc, r, "Space production cost (expect $10,000)", f"={f('C_S_prod')}/({f('p_S')}*{ySpre})", USD2); r += 1
    chk_Cdel = put(wc, r, "C_S_del (expect $350,000)", f"={f('C_S_prod')}+{f('L')}*{f('M')}+{f('C_return')}+{f('C_other')}", USD); r += 1
    chk_cS = put(wc, r, "Space delivered cost (expect $21,875)", f"={chk_Cdel}/({f('p_S')}*{yS})", USD2); r += 1
    put(wc, r, "Savings (expect -118.75%)", f"=1-{chk_cS}/{chk_cE}", "0.00%"); r += 1
    put(wc, r, "r_g_min (expect 4.375)", f"={chk_Cdel}*{f('U')}/({chk_cE}*{f('p_S')}*{f('N0')}*{f('g_E')}*{f('q_E')}*{f('r_q')}*{f('s')})", "0.000"); r += 1
    put(wc, r, "L_max (expect -$1,800/kg)", f"=({chk_cE}*{f('p_S')}*{yS}-{f('C_S_prod')}-{f('C_return')}-{f('C_other')})/{f('M')}", USD); r += 1
    put(wc, r, "P(r_g > 4.375) for triangular(1,2,6) (expect 13.203125%)", "=(6-4.375)^2/((6-1)*(6-2))", "0.000000%"); r += 1


def build_workbook(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    lines = [
        ("Microeconomics of Stem-Cell Manufacturing: Earth vs Low Earth Orbit", TITLE),
        ("Question: What is the first product that will be fundamentally cheaper to make in microgravity, and at what launch, manufacturing and return costs is it viable?", BOLD),
        ("Built 24 Oct 2026. Currency: USD, price year 2026 (source prices from other years are NOT inflation-adjusted).", BLACK),
        ("Three scenarios, same formulas: CMV_* = SPACE VIABILITY CASE (heart muscle cells on an autonomous free-flyer with projected launch/return prices; every bet flagged). CM_* = heart muscle cells with today's ISS logistics. HSC_* = cord-blood blood stem cells.", BOLD),
        ("The calculator reads these same values (app/scenario_data.js is generated by build_spreadsheet.py).", BLACK),
        ("", BLACK),
        ("How to use", BOLD),
        ("1. Edit BLUE cells on *_Assumptions (Low / Most likely / High). Black cells are formulas; green cells link to another sheet.", BLACK),
        ("2. *_Earth and *_Space build cost per batch and good cells per batch from the 'Most likely' column.", BLACK),
        ("3. *_Comparison answers the question: cost per unit, and the maximum launch price, return cost and orbital manufacturing cost at which space still matches Earth.", BLACK),
        ("4. *_Probabilities lists key claims with the author's subjective probability they are true. Discussion aids only; they do not feed the model.", BLACK),
        ("5. PRD_Check reproduces the specification's worked example to prove the formulas.", BLACK),
        ("6. To push changes into the calculator, edit build_spreadsheet.py and rerun it (editing the xlsx alone does not update the app).", BLACK),
        ("", BLACK),
        ("Evidence honesty", BOLD),
        ("Evidence status: measured in target system (Strong) / measured in another system (Moderate) / engineering estimate (Weak) / expert judgment or speculative (Very weak).", BLACK),
        ("Source verification: READ = full source page opened; SNIPPET = value seen only in a search summary; verify before citing.", BLACK),
        ("Key finding: Earth's best 3D bioreactors already reach ~88-94% correct heart cells, so a pure purity gain is capped near 1.06-1.14x. The thesis needs MORE correct cells per input cell (r_g), which is supported only by simulated-microgravity data so far.", BLACK),
    ]
    for i, (t, fnt) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=t)
        c.font, c.alignment = fnt, Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 150

    for pre in ("CMV", "CM", "HSC"):
        build_scenario(wb, pre, SCEN[pre])
    build_prd_check(wb)

    wsrc = wb.create_sheet("Sources")
    wsrc["A1"], wsrc["A1"].font = "Sources (inventory, not votes). Verification: READ = page opened; SNIPPET = search summary only.", TITLE
    header(wsrc, 3, ["ID", "Title", "Year", "URL", "Verification", "System", "Side", "Extracted finding", "Direction", "Linked inputs"],
           [6, 60, 8, 50, 11, 30, 7, 70, 20, 18])
    for i, src in enumerate(SOURCES, 4):
        for col, val in enumerate(src, 1):
            c = wsrc.cell(row=i, column=col, value=val)
            c.font, c.alignment = BLACK, WRAP

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def scenario_js(sc):
    params, evid = {}, {}
    for p in sc["P"]:
        if not p["in_engine"]:
            continue
        if p["dist"] == "fixed" or p["low"] == p["high"]:
            params[p["id"]] = {"value": p["mode"], "dist": {"type": "fixed"}}
        else:
            params[p["id"]] = {"value": p["mode"], "dist": {"type": "triangular", "min": p["low"], "mode": p["mode"], "max": p["high"]}}
        e = {"status": p["status"], "sources": p["sources"], "basis": p["basis"], "label": p["label"],
             "bet": p["limits"].startswith("BET")}
        if p["status"] not in ("definition",) and not p["status"].startswith("not extracted"):
            e["low"], e["high"] = p["low"], p["high"]
        evid[p["id"]] = e
    expected = model(point_values(sc["P"]))
    rg = sc["by"]["r_g"]
    expected["p_rg_exceeds_breakeven"] = tri_exceed(expected["r_g_min"], rg["low"], rg["mode"], rg["high"])
    return dict(params=params, paramEvidence=evid, expected=expected, product={"title": sc["title"], "unit": sc["unit"]})


def write_js(path):
    sources = [dict(id=s[0], title=s[1], year=s[2], url=s[3], verification=s[4], cell=s[5], finding=s[7], direction=s[8], params=s[9],
                    limits=("value verified on source page" if s[4] == "READ" else "value from search summary only; verify before citing"))
               for s in SOURCES]
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("// GENERATED by build_spreadsheet.py from the same tables as Stem_Cell_Microeconomics.xlsx. Do not edit by hand.\n")
        for pre in ("CMV", "CM", "HSC"):
            d = scenario_js(SCEN[pre])
            d["sources"] = sources if pre == "HSC" else []
            fh.write(f"self.{SCEN[pre]['js']} = " + json.dumps(d, indent=1) + ";\n")
    return {pre: scenario_js(SCEN[pre])["expected"] for pre in SCEN}


if __name__ == "__main__":
    build_workbook(os.path.join(HERE, "Stem_Cell_Microeconomics.xlsx"))
    exp = write_js(os.path.join(HERE, "app", "scenario_data.js"))
    print("Wrote Stem_Cell_Microeconomics.xlsx and app/scenario_data.js")
    for pre, e in exp.items():
        print(pre)
        for k in ("C_E_del", "C_S_del", "Q_E", "Q_S", "c_E", "c_S", "savings", "L_max", "C_return_max", "C_S_prod_max", "r_g_min", "r_q_min", "r_q_cap", "p_rg_exceeds_breakeven"):
            print(f"  {k:24s} {e[k]}")

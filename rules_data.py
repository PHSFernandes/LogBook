# -*- coding: utf-8 -*-
"""
Catálogo oficial de Subcompetências e Regras de Mínimos (MIMV - Egas Moniz).
Inclui parser de blocos combinados de espécies (ex: "D/C + R/ExB/ExRo/ExO/F" com mínimos "3+3")
e flags de exceção definidos pela coordenação.
"""

SPECIES_MAP = {
    "D": "Dog",
    "C": "Cat",
    "R": "Rabbit",
    "ExB": "Exotic - Bird",
    "ExRo": "Exotic - Rodents",
    "ExO": "Exotic - Others",
    "F": "Fish",
    "E": "Equine",
    "B": "Bovine",
    "O": "Ovine",
    "S": "Swine",
    "G": "Goat",
    "P": "Poultry",
}

ALL_SPECIES = [
    "Equine", "Dog", "Swine", "Cat", "Exotic - Bird", "Ovine",
    "Bovine", "Goat", "Exotic - Others", "Exotic - Rodents",
    "Poultry", "Rabbit", "Fish"
]

LIVESTOCK_SPECIES = ["Swine", "Ovine", "Bovine", "Goat", "Poultry"]
COMPANION_SPECIES = ["Dog", "Cat", "Rabbit", "Exotic - Bird", "Exotic - Rodents", "Exotic - Others", "Fish"]
EQUINE_SPECIES = ["Equine"]

# (Código, Descrição, Regra Espécie, Regra Mínimo)
RAW_COMPETENCE_RULES = [
    ("1-Ha01", "Approaching the animal correctly and being aware of the rejection reaction", "D/C + R/ExB/ExRo/ExO/F", "3+3"),
    ("1-Ha02", "Placing and removing animals from the cage/transport box", "D/C + R/ExB/ExRo/ExO/F", "3+3"),
    ("1-Ha03", "Fitting the restraining loop/muzzle", "D", "3"),
    ("1-Ha04", "Putting on an animal collar (Elizabeth collar)", "D/C", "3"),
    ("1-Ha05", "Restraining the animal (clinical examination, blood/urine collection, medication, treatments, dressings)", "D + C + R/ExB/ExRo/ExO/F", "3+3+3"),
    ("1-Ha06", "Approaching the farm animal correctly and being aware of the rejection reaction", "E + B/O/G", "3+3"),
    ("1-Ha07", "Fitting a bridle and leading the horse by hand (walk and trot)", "E", "3"),
    ("1-Ha08", "Restraining the animal (harness/ bridle/ ear torsion/ limb flexion/etc)", "Any", "3"),
    ("1-An01", "Individual animal (Anamnesis)", "Any", "3"),
    ("1-An02", "Group of animals (Anamnesis)", "E/B/G/O/S", "3"),
    ("1-GE01", "Identify most common breeds / most common coat", "Any", "3"),
    ("1-GE02", "Determine the sex of the animal", "Any", "3"),
    ("1-GE03", "Determine approximate age of animal", "Any", "3"),
    ("1-GE04", "Rectal temperature", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-GE05", "Heart (HR) and respiratory rate (RR)", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-GE06", "Mucosal coloration and capillary refill time (CRT)", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-GE07", "Arterial pulse", "Any", "3"),
    ("1-GE08", "Pain assessment", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-GE09", "Degree of dehydration", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+2"),
    ("1-GE10", "Nutrition / Assess body condition", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+2"),
    ("1-GE11", "Animal welfare / Pain evaluation", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+2"),
    ("1-Ly01", "Location and palpation of the main peripheral lymph nodes", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-Re01", "Percussion and thoracic auscultation", "Any", "3"),
    ("1-CV01", "Identification and assessment of heart sounds", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-GT01", "Examination of the oral cavity and teeth", "Any", "3"),
    ("1-GT02", "Abdominal palpation and/or auscultation", "Any", "3"),
    ("1-GT03", "Rectal examination to assess the digestive tract", "E/B", "2"),
    ("1-De01", "Recognizing primary and secondary dermatological lesions", "Any", "3"),
    ("1-De02", "Observation of the ear canal with an otoscope", "Any", "1"),
    ("1-De03", "Wood's lamp examination", "Any", "3"),
    ("1-Ey01", "Perform direct and indirect ophthalmoscopic examination", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ey02", "Evaluation of the threat reflex", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ey03", "Pupillary reflex to light", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ey04", "Schirmer test", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ey05", "Intraocular pressure measurement with tonopen", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ey06", "Fluorescein staining test", "D/C/R/ExB/ExRo/ExO/F/E", "2"),
    ("1-Ort01", "Identification and assessment of the most important joints and tendons", "D/C/E", "3"),
    ("1-Ort02", "Static examination", "D/C/E", "3"),
    ("1-Ort03", "Dynamic examination", "D/C/E", "3"),
    ("1-Ort04", "Limb pain assessment", "D/C/E", "3"),
    ("1-Ort05", "Regional nerve blocks", "E", "3"),
    ("1-Ho01", "Removing the shoe", "E", "3"),
    ("1-Ho02", "Using hoof clamps", "E", "3"),
    ("1-Ho03", "Trimming the hoof", "E", "3"),
    ("1-Ne01", "Patellar reflex", "D/C", "3"),
    ("1-Ne02", "Withdrawal reflex", "D/C", "3"),
    ("1-Ne03", "Pannicular reflex", "D/C/E", "3"),
    ("1-Ne04", "Proprioception assessment", "D/C/E", "3"),
    ("1-Ne05", "Deep pain assessment", "D/C/E", "3"),
    ("1-Ne06", "Cranial nerve examination", "D/C/E", "3"),
    ("1-Ne07", "Spinal reflex examination", "D/C/E", "3"),
    ("1-Ur01", "Palpation and assessment of testicles and penis", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-Ur02", "Rectal examination (prostate evaluation)", "D", "3"),
    ("1-Ur03", "Palpation of the cervix and evaluation of the vagina", "D/C", "3"),
    ("1-Ur04", "Rectal palpation for reproductive tract assessment/pregnancy diagnosis", "B + E", "3+3"),
    ("1-Ur05", "Palpation of the mammary gland", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("1-Ur06", "Preparing and assessing semen", "B/G/O/S/E/C", "3"),
    ("1-Ur07", "Collecting and interpreting vaginal cytology", "D", "3"),
    ("1-Ur08", "Identifying and correcting dystocia", "B + E", "3+3"),
    ("2-Li01", "Make a list of problems in clinical cases", "Any", "3"),
    ("2-Li02", "Make a list of differential diagnoses in clinical cases", "Any", "3"),
    ("2-Li03", "Choosing the appropriate diagnostic tests in clinical cases", "Any", "3"),
    ("2-Li04", "Follow case during treatment", "Any", "3"),
    ("2-Li05", "Discuss case progression with clients (owner concern, financial, empathy)", "Any", "3"),
    ("2-Li06", "Give discharge instructions in a case", "Any", "3"),
    ("2-Li07", "Follow up by telephone of cases to assess progression/end of problem", "Any", "3"),
    ("2-Li08", "Participate in primary care cases (vaccination etc)", "Any", "3"),
    ("2-Li09", "Participate in clinical referral cases", "Any", "3"),
    ("2-Bl01", "Blood sample collection - Jugular vein", "Any", "3"),
    ("2-Bl02", "Blood sample collection - Cephalic/saphenous vein", "Any", "3"),
    ("2-Bl03", "Blood sample collection - Tail vein", "B", "3"),
    ("2-UC01", "Free collection/compression of the bladder by trans-abdominal palpation", "Any", "3"),
    ("2-UC02", "Urinary catheterization", "Any", "3"),
    ("2-UC03", "Cystocentesis", "D/C", "3"),
    ("2-Sk01", "Skin scraping - superficial/deep", "Any", "3"),
    ("2-Sk02", "Cytological skin sampling - appositional/tape/swab/brush", "Any", "3"),
    ("2-Sk03", "Skin biopsy (Punch, Excisional)", "Any", "2"),
    ("2-TF01", "Fine needle aspiration (FNA) / Biopsy (Punch, Excisional)", "Any", "3"),
    ("2-TF02", "Swab for microbiological analysis", "Any", "3"),
    ("2-TF03", "Stool sample", "Any", "3"),
    ("2-TF04", "Cloaca sample", "ExB/ExO/P", "3"),
    ("2-TF05", "Milk sample (e.g. mastitis, antibody)", "B/G/O", "3"),
    ("2-Sa01", "Selection of collection tubes/packaging - Blood and body fluids", "Any", "3"),
    ("2-Sa02", "Selection of collection tubes/packaging - Skin/tissue", "Any", "3"),
    ("2-Sa03", "Selection of collection tubes/packaging - Microbiology", "Any", "3"),
    ("2-La01", "Preparing and staining a blood smear", "Any", "3"),
    ("2-La02", "Reading a hematocrit", "Any", "3"),
    ("2-La03", "Performing a differential white cell count", "Any", "3"),
    ("2-La04", "Performing biochemical tests", "Any", "3"),
    ("2-Ur01", "Urinary specific gravity", "Any", "3"),
    ("2-Ur02", "Rapid strips (urine)", "Any", "3"),
    ("2-Ur03", "Sediment examination (type II urine)", "Any", "3"),
    ("2-MO01", "Microscopic examination of skin samples - Ectoparasites", "Any", "3"),
    ("2-MO02", "Microscopic examination of skin samples - Dermatophytes/fungi", "Any", "3"),
    ("2-MO03", "Microscopic examination of skin samples - Bacteria", "Any", "3"),
    ("2-Fc01", "Coprological flotation analysis", "Any", "3"),
    ("2-Fc02", "Coprological sedimentation analysis", "Any", "3"),
    ("2-Fc03", "Examination of fresh and stained fecal smears", "Any", "3"),
    ("2-Mb01", "Preparing and staining bacterial preparations for microscopic examination", "Any", "3"),
    ("2-Mb02", "Propagating samples on suitable media", "Any", "3"),
    ("2-Mb03", "Identification of the pathogen by API or other biochemical test", "Any", "3"),
    ("2-Mb04", "Performing and evaluating antibiograms", "Any", "3"),
    ("2-Mb05", "Performing and interpreting ELISA test results", "Any", "3"),
    ("2-Mb06", "Performing and interpreting a Californian test on milk samples", "Any", "3"),
    ("2-Mb07", "Tuberculinization test", "Any", "3"),
    ("2-Rx01", "Radiological protection (gloves, apron and thyroid protector)", "Any", "3"),
    ("2-Rx02", "Positioning for X-ray - Appendicular skeleton", "Any", "3"),
    ("2-Rx03", "Positioning for X-ray - Axial skeleton", "Any", "3"),
    ("2-Rx04", "Positioning for X-ray - Thorax", "Any", "3"),
    ("2-Rx05", "Positioning for X-ray - Abdomen", "Any", "3"),
    ("2-Rx06", "Selection of radiographic constants and appropriate cassette", "Any", "3"),
    ("2-Rx07", "Choosing appropriate contrast medium (barium sulphate/iodized contrast)", "Any", "3"),
    ("2-Rx08", "Evaluating and interpreting radiographic signs", "Any", "3"),
    ("2-Rx09", "Identifying artifacts and correcting them", "Any", "3"),
    ("2-US01", "Preparing animals for ultrasound examination and witnessing US exams", "Any", "3"),
    ("2-US02", "Echographic examination training on a phantom (Skill LAB)", "EXCLUDED", "0"),
    ("2-ECG", "Performing an electrocardiogram and evaluating ECG tracings", "Any", "1"),
    ("2-END", "Attending and interpreting endoscopy findings", "Any", "1"),
    ("2-CT", "Attending and interpreting CT scan findings", "Any", "1"),
    ("2-TT01", "Plan a surgical treatment", "Any", "3"),
    ("2-TT02", "Plan a Medical treatment", "Any", "3"),
    ("2-TT03", "Registration of Clinical Case report", "Any", "3"),
    ("2-Nc01", "Performing a necropsy", "D/C/E/R/ExB/ExRo/ExO/F/S + B/O/G + R/ExB/ExRo/ExO/F/B/P", "3+1+1"),
    ("2-Nc02", "Macroscopic assessment of normal and pathological changes in tissues and organs", "Any", "3"),
    ("2-Nc03", "Taking and packaging samples for complementary tests", "Any", "3"),
    ("2-Nc04", "Drawing up a necropsy report", "Any", "3"),
    ("3-Em01", "Recognizing life-threatening situations and initiating appropriate treatment", "D/C/R/ExB/ExRo/ExO/F + B/E/S/G/O/P/R", "3+2"),
    ("3-Em02", "Participation on basic cardiopulmonary resuscitation / neonate resuscitation / O2 therapy", "Any", "1"),
    ("3-Em03", "Participate thoracentesis / Participate in a tracheostomy", "Any", "1"),
    ("3-Em04", "Naso/oro-gastric tube placement", "Any", "1"),
    ("3-Em05", "Participate abdominocentesis / Participate on a trocarization", "Any", "1"),
    ("3-Ad01", "Oral administration of drugs - Tablets", "D/C", "3"),
    ("3-Ad02", "Oral administration of drugs - Paste", "B/G/O/E", "3"),
    ("3-Ad03", "Parenteral drug administration - subcutaneous", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("3-Ad04", "Parenteral drug administration - intramuscular", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("3-Ad05", "Parenteral drug administration - intravenous", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("3-Ad06", "Parenteral drug administration - intra-dermal/topical", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+3"),
    ("3-Ad07", "Fluid therapy - Determining the volume and rate of administration", "Any", "3"),
    ("3-Ad08", "Fluid therapy - Placement of intravenous catheter and venoclysis system", "Any", "3"),
    ("3-Ad09", "Fluid therapy - Monitoring the degree of hydration", "Any", "3"),
    ("3-An01", "Realization of sedation", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+2"),
    ("3-An02", "Choosing an anesthetic/analgesic protocol", "Any", "3"),
    ("3-An03", "Selecting and preparing anesthesia equipment", "Any", "3"),
    ("3-An04", "Administer premedication and induction agent", "Any", "3"),
    ("3-An05", "Placing an endotracheal tube", "Any", "2"),
    ("3-An06", "Placement of pulse oximeter", "Any", "3"),
    ("3-An07", "Maintenance and monitoring of anesthesia / Perform manual IPPV", "D/C/R/ExB/ExRo/ExO/F + E/B/O/S/G", "3+1"),
    ("3-An09", "Local anesthesia - Infiltration / Block / Epidural", "Any", "1"),
    ("3-An10", "Local anesthesia - Topical application", "Any", "1"),
    ("3-Pr01", "Calculating the correct dose of the active ingredient", "Any", "3"),
    ("3-Pr02", "Prescribing a simple prescription", "Any", "3"),
    ("3-Pr03", "Prescribing a medicated food", "Any", "3"),
    ("3-Pr04", "Prescribing a master prescription", "Any", "3"),
    ("3-Pn01", "Recognizing, scoring and managing pain during and after surgery", "Any", "3"),
    ("3-Pn02", "Chronic pain management", "Any", "3"),
    ("3-Su01", "Preparing and sterilizing surgical instruments and equipment", "PENDING_MIN", "0"),
    ("3-Su02", "Surgical preparation as a surgeon's assistant", "D/C/R/ExB/ExRo/ExO/F/E", "3"),
    ("3-Su03", "Preparing the animal and the operating field (trichotomy/asepsis)", "D/C/R/ExB/ExRo/ExO/F/E", "3"),
    ("3-Su04", "Positioning the patient on the operating table", "D/C/R/ExB/ExRo/ExO/F/E", "3"),
    ("3-Su05", "Opening the surgical packs and layout of the surgical kit", "D/C/R/ExB/ExRo/ExO/F/E", "3"),
    ("3-Su06", "Wound debridement", "Any", "3"),
    ("3-Su07", "Abscess drainage", "Any", "3"),
    ("3-Su08", "Skin incision", "Any", "3"),
    ("3-Su09", "Participate in a placement of a surgical drain", "Any", "3"),
    ("3-Su10", "Participate in a removal of a simple mass on the skin", "Any", "3"),
    ("3-Su11", "Hemostasis in superficial bleeding", "Any", "3"),
    ("3-Su12", "Selecting appropriate suture materials", "EXCLUDED", "0"),
    ("3-Su13", "Performing simple skin suturing", "EXCLUDED", "0"),
    ("3-Su14", "Perform simple continuous suturing", "EXCLUDED", "0"),
    ("3-Su15", "Perform vertical/horizontal U-sutures", "EXCLUDED", "0"),
    ("3-Su16", "Perform intestinal suturing (cushing/lambert)", "EXCLUDED", "0"),
    ("3-Su17", "Performing a surgical knot", "EXCLUDED", "0"),
    ("3-Su18", "Demonstrate correct use of surgical equipment", "Any", "3"),
    ("3-Su19", "Ovariohysterectomy", "D/C", "3"),
    ("3-Su20", "Orchiectomy", "D/C", "3"),
    ("3-Su21", "Laparotomy", "D/C", "3"),
    ("3-Su22", "Participate in an equine orchiectomy", "E", "3"),
    ("3-Su23", "Participate in 3 referral surgeries (including an orthopedic surgery)", "D/C/R/ExB/ExRo/ExO/F", "3"),
    ("3-Su24", "Participate in 3 soft tissue surgeries", "Any", "3"),
    ("3-Su25", "Umbilical hernia repair", "Any", "1"),
    ("3-Su26", "Foreign body removal (ear, skin, eye, nose, etc)", "Any", "3"),
    ("3-Su27", "Post-surgical monitoring of the animal", "Any", "3"),
    ("3-Su28", "Post-surgical counseling at discharge (healing, feeding, rehabilitation)", "Any", "3"),
    ("3-Su29", "Dehorning", "B", "1"),
    ("3-De01", "Examination of teeth and identification of lesions", "Any", "3"),
    ("3-De02", "Tartar removal and dental polishing", "Any", "1"),
    ("3-Eu01", "Participating in a euthanasia and explaining indications and methods", "Any", "3"),
    ("3-Eu02", "Correct handling of animals and equipment/product used for euthanasia", "Any", "3"),
    ("3-Eu03", "Dispose of corps appropriately, including biosafety rules", "Any", "3"),
    ("3-Mn01", "Proper environment and housing conditions", "Any", "3"),
    ("3-Mn02", "Disease prevention", "Any", "3"),
    ("3-Mn03", "Feeding and nutritional needs", "Any", "3"),
    ("3-Mn04", "Vaccination protocol", "Any", "3"),
    ("3-Mn05", "Internal and external deworming protocol", "Any", "3"),
    ("3-Mn06", "Identification and evaluation of feed for production animals", "Any", "3"),
    ("3-Mn07", "Advice on specific dietary needs - nutritional deficiencies", "Any", "3"),
    ("3-Mn08", "Advice on specific dietary needs - stages of life", "Any", "3"),
    ("3-Mn09", "Advice on specific dietary needs - specific conditions", "Any", "3"),
    ("3-Mn10", "Advise owners on the species/breed to choose", "Any", "3"),
    ("3-Mn11", "Advise owners on socialization, behaviour and training", "Any", "3"),
    ("3-Mn12", "Educate owners on reproductive management", "Any", "3"),
    ("3-Mn13", "Inform owners of relevant legislation", "Any", "3"),
    ("3-Mn14", "Evaluate and implement a farm's production, reproductive and health records", "Any", "3"),
    ("3-Cer01", "Ante-post-mortem inspection of production animals", "Any", "3"),
    ("3-Cer02", "Equine passport", "E", "3"),
    ("3-Cer03", "Passport for pets", "Any", "3"),
    ("3-Cer04", "Emergency slaughter/ culling", "Any", "1"),
    ("3-Cer05", "Know how to act in cases of - Notifiable diseases", "Any", "1"),
    ("3-Cer06", "Know how to act in cases of - Reportable diseases", "Any", "1"),
    ("3-Cer07", "Know how to act in cases of - Zoonotic diseases", "Any", "1"),
    ("3-Cer08", "Correct and safe administration of medication", "Any", "3"),
    ("3-Cer09", "Side effects of medication", "Any", "3"),
    ("3-Cer10", "Safety interval of medication", "Any", "3"),
    ("3-Cer11", "Correct storage and handling of medication", "Any", "3"),
    ("3-Nur01", "Execute and remove a simple dressing", "Any", "3"),
    ("3-Nur02", "Execute and remove a hard dressing (e.g. Robert Jones)", "Any", "2"),
    ("3-Nur03", "Ear cleaning", "Any", "3"),
    ("3-Nur04", "Hoof cleaning", "E/B", "3"),
    ("3-Nur05", "Trimming nails and claws", "Any", "3"),
    ("3-Nur06", "Beak trimming", "ExB/P", "3"),
    ("3-Nur07", "De-maggoting", "Any", "3"),
]


def parse_rule_blocks(species_rule: str, min_rule: str):
    """
    Converte regras como ('D/C + R/ExB/ExRo/ExO/F', '3+3') em lista de blocos:
    [
        {'label': 'D/C', 'species': ['Dog', 'Cat'], 'min': 3},
        {'label': 'R/ExB/ExRo/ExO/F', 'species': ['Rabbit', ...], 'min': 3}
    ]
    """
    sp_clean = str(species_rule).strip()
    if sp_clean in ("EXCLUDED", "PENDING_MIN", ""):
        return []

    sp_parts = [p.strip() for p in sp_clean.split("+")]
    min_parts = [int(m.strip()) for m in str(min_rule).split("+") if m.strip().isdigit()]

    if len(min_parts) < len(sp_parts):
        min_parts = min_parts + [min_parts[-1] if min_parts else 0] * (len(sp_parts) - len(min_parts))

    blocks = []
    for sp_grp, req_min in zip(sp_parts, min_parts):
        if sp_grp.lower() == "any":
            sp_list = list(ALL_SPECIES)
        else:
            tokens = [t.strip() for t in sp_grp.split("/") if t.strip()]
            sp_list = [SPECIES_MAP[t] for t in tokens if t in SPECIES_MAP]
        blocks.append({
            "label": sp_grp,
            "species": sp_list,
            "min": req_min
        })
    return blocks


def get_default_rules_list():
    """Retorna lista estruturada de regras para edição e cálculo."""
    items = []
    for code, desc, sp_rule, min_rule in RAW_COMPETENCE_RULES:
        if sp_rule == "EXCLUDED":
            status = "Desconsiderada"
        elif sp_rule == "PENDING_MIN":
            status = "Sem mínimo definido"
        else:
            status = "Ativa"
        items.append({
            "Código": code,
            "Subcompetência": desc,
            "Regra Espécies": sp_rule,
            "Regra Mínimo": min_rule,
            "Estado": status,
        })
    return items

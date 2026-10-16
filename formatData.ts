function main(workbook: ExcelScript.Workbook) {
  // ================= CONFIG =================
  const SOURCE_SHEET_NAME = "DATA_SHEET";
  const DEST_SHEET_NAME = "LIST_DATA";

  // Floors to include, inclusive.
  const FLOOR_MIN = 6;
  const FLOOR_MAX = 10;

  // Window/slice sizes.
  const AFTER_SLICE_LEN = 240;
  const FLOOR_SLICE_BEFORE = 160;
  const FLOOR_SLICE_AFTER = 200;

  // ================= DEBUG =================
  const DEBUG = true;

  // ================= TYPES =================
  type Cell = string | number | boolean;

  type RoomHit = {
    room: string;
    start: number;
    end: number;
  };

  // Public columns:
  // Floor, Room, Last, Dx, Svc, YN1, YN2
  type Row = [
    string,
    string,
    string,
    string,
    string,
    string,
    string
  ];

  // Internal row with patientKey at index 7.
  type RowEx = [
    string,
    string,
    string,
    string,
    string,
    string,
    string,
    string
  ];

  // ================= BASIC HELPERS =================

  // Strong normalization:
  // - Converts pipes to spaces.
  // - Removes tildes.
  // - Converts stray backslashes to spaces.
  // - Converts nonbreaking spaces.
  // - Collapses repeated whitespace.
  const normalize = (s: string): string => {
    return String(s ?? "")
      .replace(/[|]+/g, " ")
      .replace(/[~]+/g, "")
      .replace(/\\+/g, " ")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  };

  const toSpaced = (s: string): string => {
    return normalize(s);
  };

  const toTight = (s: string): string => {
    return normalize(s).replace(/\s+/g, "");
  };

  // Boilerplate/non-patient rows to ignore early.
  const isBoilerplate = (line: string): boolean => {
    return (
      /^DCN\s+SHEET\s+WITH\b/i.test(line) ||
      /^DCN\s+List\b/i.test(line) ||
      /^Unit[a]?\b/i.test(line) ||
      /Care\s*Team/i.test(line) ||
      /Encounter/i.test(line) ||
      /Flag(Name)?/i.test(line) ||
      /^Page\s+\d+\s+of\s+\d+/i.test(line) ||
      /^Printed\s+by\b/i.test(line) ||
      /^Patients:\s*\d+/i.test(line) ||
      /\bCH\s+PAVA\s+OR\b/i.test(line) ||
      /\bOR\s*\/-\b/i.test(line) ||

      // Ignore a line that is only CH, optionally with small OCR noise.
      /^\s*CH\s*$/i.test(line) ||
      /^\s*CH\s*[~\-\|\(\)]*\s*$/i.test(line)
    );
  };

  // If the combined window contains a noisy "CH ~~" fragment,
  // skip the window entirely.
  const hasNoisyCH = (text: string): boolean => {
    return /\bCH\s*~+/i.test(text);
  };

  // Matches a token with optional spaces between its characters.
  //
  // Example:
  // tokenPattern("CHPAVATT1") matches:
  // CHPAVATT1
  // CH PAVA TT1
  const tokenPattern = (s: string): RegExp => {
    return new RegExp(
      s
        .split("")
        .map((ch: string): string => {
          const esc = ch.replace(
            /[-\/\\^$*+?.()|[\]{}]/g,
            "\\$&"
          );

          return esc + "\\s*";
        })
        .join(""),
      "gi"
    );
  };

  // Promote a trailing "<floor>T1" or "<floor>T2" token found later
  // in the window.
  //
  // Example:
  // CH PAVA 225/225A ... 9T2 UNI ...
  //
  // Becomes:
  // CH PAVA 9 T2 225/225A ... UNI ...
  const promoteFloorToken = (s: string): string => {
    const re: RegExp =
      /(\bCH\s*PAVA\b)([^\n]*?)(\b\d{3}\s*\/\s*\d{3}[A-Z0-9]\b)([\s\S]*?\b(\d{1,2})\s*T([12])\b)/i;

    const m: RegExpExecArray | null = re.exec(s);

    if (!m) {
      return s;
    }

    const pava: string = m[1];
    const mid: string = m[2];
    const bed: string = m[3];
    const tail: string = m[4];
    const floor: string = m[5];
    const towerNumber: string = m[6];

    const floorToken = new RegExp(
      `\\b${floor}\\s*T${towerNumber}\\b`,
      "i"
    );

    const newTail: string = tail
      .replace(floorToken, "")
      .replace(/\s{2,}/g, " ")
      .trim();

    const promoted: string =
      `${pava} ${floor} T${towerNumber} ${mid}${bed} ${newTail}`
        .replace(/\s{2,}/g, " ")
        .trim();

    const before: string = s.slice(0, m.index);
    const after: string = s.slice(m.index + m[0].length);

    return (before + promoted + after)
      .replace(/\s{2,}/g, " ")
      .trim();
  };

  // ================= SPACED OCR FIXES =================

  const applyOcrTokenFixesSpaced = (input: string): string => {
    let sp: string = input;

    // IMPORTANT FIX:
    // Ignore an erroneous OCR "T" inserted between PAVA and the floor.
    //
    // Examples:
    // CHPAVAT7 T1/127/127A -> CH PAVA 7 T1/127/127A
    // CH PAVA T 7 T1      -> CH PAVA 7 T1
    //
    // This supports floors 6 through 10.
    sp = sp.replace(
      /\bCH\s*PAVA\s*T\s*(10|[6-9])\b/gi,
      "CH PAVA $1"
    );

    // Explicit OCR token mappings.
    sp = sp.replace(
      tokenPattern("CHPAVAG6T2"),
      "CH PAVA 6 T2"
    );

    sp = sp.replace(
      tokenPattern("CHPAVAGBT2"),
      "CH PAVA 6 T2"
    );

    // Heal CG -> CH OCR swap.
    sp = sp.replace(
      /\bCG\s*PAVA\b/gi,
      "CH PAVA"
    );

    // Ensure a space between the floor and tower token.
    //
    // CH PAVA 10T2 -> CH PAVA 10 T2
    sp = sp.replace(
      /\bCH\s*PAVA\s*(\d{1,2})T([12])\b/gi,
      "CH PAVA $1 T$2"
    );

    // If T1/T2 is immediately followed by a room digit,
    // force a space before the room.
    //
    // CH PAVA 10T2216/216A
    // becomes
    // CH PAVA 10 T2 216/216A
    sp = sp.replace(
      /(\bCH\s*PAVA\b[^\n]{0,12}?T[12])(?=\s*\d)/gi,
      "$1 "
    );

    // Short-left bed fix:
    // 16/216A -> 216/216A
    sp = sp.replace(
      /\b(\d{1,2})\s*\/\s*(\d{3})([A-Z0-9])\b/g,
      (
        _match: string,
        _short: string,
        right: string,
        bedMarker: string
      ): string => {
        return `${right}/${right}${bedMarker}`;
      }
    );

    // No-slash long-digit blob directly after the PAVA T# header.
    //
    // 2301230A -> 230/230A
    // 2320232A -> 232/232A
    sp = sp.replace(
      /(\bCH\s*PAVA\b[^\n]{0,16}?T[12])\s*(\d{6,8})([A-Z0-9])\b/gi,
      (
        _match: string,
        head: string,
        digits: string,
        bedMarker: string
      ): string => {
        const right: string = digits.slice(-3);
        return `${head} ${right}/${right}${bedMarker}`;
      }
    );

    // Canonicalize separated or malformed six-digit bed formats.
    //
    // 2 12)2 12A -> 212/212A
    sp = sp.replace(
      /\b(\d)\s*(\d)\s*(\d)\s*[/\)\|]?\s*(\d)\s*(\d)\s*(\d)\s*([A-Z0-9])\b/g,
      (
        _match: string,
        a: string,
        b: string,
        c: string,
        d: string,
        e: string,
        f: string,
        bedMarker: string
      ): string => {
        return `${a}${b}${c}/${d}${e}${f}${bedMarker}`;
      }
    );

    // 211)211A -> 211/211A
    sp = sp.replace(
      /\b(\d{3})\)\s*(\d{3})([A-Z0-9])\b/g,
      "$1/$2$3"
    );

    // 222220A -> 222/220A
    sp = sp.replace(
      /\b(\d{3})(\d{3})([A-Z0-9])\b/g,
      "$1/$2$3"
    );

    // 2121212A -> 212/212A
    sp = sp.replace(
      /\b(\d{3})\d\1([A-Z0-9])\b/g,
      "$1/$1$2"
    );

    // Existing OCR floor mappings.
    sp = sp.replace(
      tokenPattern("CHPAVAGT2"),
      "CH PAVA 6 T2"
    );

    sp = sp.replace(
      tokenPattern("CHPAVATT1"),
      "CH PAVA 7 T1"
    );

    sp = sp.replace(
      tokenPattern("CHPAVATT!"),
      "CH PAVA 7 T1"
    );

    sp = sp.replace(
      tokenPattern("CHPAVATT2"),
      "CH PAVA 7 T2"
    );

    sp = sp.replace(
      tokenPattern("CHPAVA10T22"),
      "CH PAVA 10 T2"
    );

    sp = sp.replace(
      tokenPattern("CHPAVA10TT2"),
      "CH PAVA 10 T2 "
    );

    // Normalize CH PAVA spacing.
    sp = sp.replace(
      /\bCG\s*PAVA\b/gi,
      "CH PAVA"
    );

    sp = sp.replace(
      /\bCH\s*PAVA\b/gi,
      "CH PAVA"
    );

    // Common OCR artifacts.
    sp = sp.replace(
      /\bProgressi\b/gi,
      "Progressive"
    );

    sp = sp.replace(
      /\bve\b/gi,
      ""
    );

    sp = sp.replace(
      /\[\s*/g,
      " "
    );

    sp = sp.replace(
      /\bGII?HPB\b/gi,
      "GIHPB"
    );

    sp = sp.replace(
      /\bIcy\b/gi,
      "ICU"
    );

    sp = sp.replace(
      /\bIcU\b/gi,
      "ICU"
    );

    sp = sp.replace(
      /\bA\s*contact\b/gi,
      "Contact"
    );

    sp = sp.replace(
      /\bAcontactp\b/gi,
      "Contact"
    );

    // Q mistaken for 9 immediately before T1/T2.
    sp = sp.replace(
      /\bCH\s*PAVA\s*Q\s*(T[12])\b/gi,
      "CH PAVA 9 $1"
    );

    // Additional PAVA/T spacing.
    sp = sp.replace(
      /(\bCH\s*PAVA\b[^\n]{0,12}?T[12])(?=\s*\d)/gi,
      "$1 "
    );

    // Re-run robust bed canonicalization after the PAVA fixes.
    sp = sp.replace(
      /\b(\d)\s*(\d)\s*(\d)\s*[/\)\|]?\s*(\d)\s*(\d)\s*(\d)\s*([A-Z0-9])\b/g,
      (
        _match: string,
        a: string,
        b: string,
        c: string,
        d: string,
        e: string,
        f: string,
        bedMarker: string
      ): string => {
        return `${a}${b}${c}/${d}${e}${f}${bedMarker}`;
      }
    );

    sp = sp.replace(
      /\b(\d{3})\)\s*(\d{3})([A-Z0-9])\b/g,
      "$1/$2$3"
    );

    sp = sp.replace(
      /\b(\d{3})(\d{3})([A-Z0-9])\b/g,
      "$1/$2$3"
    );

    sp = sp.replace(
      /\b(\d{3})\d\1([A-Z0-9])\b/g,
      "$1/$1$2"
    );

    // T1/T2 followed by OCR punctuation.
    sp = sp.replace(
      /((?:CH\s*)?PAVA?[^\n]{0,24}?T[12])[)\|?]/gi,
      "$1 "
    );

    // Rare OCR T used for the last digit of the room.
    //
    // 221/22TA -> 221/221A
    sp = sp.replace(
      /\b(\d{3})\s*\/\s*(\d{2})T([A-Z0-9])\b/g,
      (
        original: string,
        left: string,
        firstTwoRightDigits: string,
        bedMarker: string
      ): string => {
        if (/^\d{3}$/.test(left)) {
          return `${left}/${firstTwoRightDigits}1${bedMarker}`;
        }

        return original;
      }
    );

    // Collapse inner spaces in right-side room digits.
    sp = sp.replace(
      /\b(\d{3})\s*\/\s*(\d)\s*(\d)\s*(\d)([A-Z0-9])\b/g,
      (
        _match: string,
        left: string,
        a: string,
        b: string,
        c: string,
        bedMarker: string
      ): string => {
        return `${left}/${a}${b}${c}${bedMarker}`;
      }
    );

    // Apply floor promotion after the basic OCR corrections.
    sp = promoteFloorToken(sp);

    return sp
      .replace(/\s{2,}/g, " ")
      .trim();
  };

  // ================= TIGHT OCR FIXES =================

  const applyOcrTokenFixesTight = (
    tightStr: string
  ): string => {
    let t: string = tightStr;

    // Heal CG -> CH in tight context.
    t = t.replace(
      /CGPAVA/gi,
      "CHPAVA"
    );

    // IMPORTANT FIX:
    // Remove an erroneous T between PAVA and floors 6 through 10.
    //
    // CHPAVAT7T1 -> CHPAVA7T1
    t = t.replace(
      /CHPAVAT(10|[6-9])/gi,
      "CHPAVA$1"
    );

    // Existing explicit mappings.
    t = t.replace(
      /CHPAVAG6T2/gi,
      "CHPAVA6T2"
    );

    t = t.replace(
      /CHPAVAGBT2/gi,
      "CHPAVA6T2"
    );

    t = t.replace(
      /CHPAVAGT2/gi,
      "CHPAVA6T2"
    );

    t = t.replace(
      /CHPAVATT1/gi,
      "CHPAVA7T1"
    );

    t = t.replace(
      /CHPAVATT!/gi,
      "CHPAVA7T1"
    );

    t = t.replace(
      /CHPAVATT2/gi,
      "CHPAVA7T2"
    );

    // Floor 10 fixes.
    t = t.replace(
      /CHPAVA10T22/gi,
      "CHPAVA10T2"
    );

    t = t.replace(
      /CHPAVA10TT2/gi,
      "CHPAVA10T2"
    );

    // Remove OCR punctuation after T1/T2.
    t = t.replace(
      /((?:CH)?PAVA?[A-Z0-9]{0,12}T[12])[)\|?]/gi,
      "$1"
    );

    return t;
  };

  // ================= FLOOR DETECTION =================

  // Floor detection around PAV/PAVA.
  //
  // It remains tolerant of T being an OCR representation of 7,
  // but CHPAVAT7 is corrected before this function runs.
  const findFloorStrict = (
    sp: string,
    tight: string
  ): string => {
    const scanAfterPAV = (s: string): string => {
      const up: string = s.toUpperCase();

      let idx: number = up.indexOf("PAV");

      if (idx < 0) {
        const chIdx: number = up.indexOf("CHPAV");

        if (chIdx >= 0) {
          idx = chIdx + 2;
        }
      }

      if (idx < 0) {
        return "";
      }

      let j: number = idx + 3;

      if (up[j] === "A") {
        j++;
      }

      const window: string = up.slice(j, j + 16);

      const mapCharToDigit = (
        ch: string
      ): string => {
        if (ch >= "0" && ch <= "9") {
          return ch;
        }

        if (ch === "T") {
          return "7";
        }

        return "";
      };

      for (let k = 0; k < window.length; k++) {
        const d1: string = mapCharToDigit(window[k]);

        if (!d1) {
          continue;
        }

        let d2: string = "";

        if (k + 1 < window.length) {
          const nextDigit: string =
            mapCharToDigit(window[k + 1]);

          if (nextDigit) {
            d2 = nextDigit;
          }
        }

        return (d1 + d2).slice(0, 2);
      }

      return "";
    };

    return scanAfterPAV(sp) || scanAfterPAV(tight) || "";
  };

  // ================= BED EXTRACTOR =================

  // Always returns the right-side three-digit room immediately
  // before the bed marker.
  const findBedWithBounds = (
    sp: string,
    _tight: string
  ): RoomHit | null => {
    // A. Normal or spaced ###/###A pattern.
    {
      const re: RegExp =
        /\b(\d)\s*(\d)\s*(\d)\s*\/\s*(\d)\s*(\d)\s*(\d)\s*([A-Z0-9])\b/g;

      const match: RegExpExecArray | null = re.exec(sp);

      if (match) {
        const start: number = match.index;
        const end: number = start + match[0].length;
        const room: string =
          `${match[4]}${match[5]}${match[6]}`;

        return {
          room,
          start,
          end
        };
      }
    }

    // B. Duplicated room patterns.
    //
    // 231231A
    // 2420242A
    // 2322327
    {
      const re: RegExp =
        /\b(\d{3})\d?\1([A-Z0-9])\b/g;

      const match: RegExpExecArray | null = re.exec(sp);

      if (match) {
        const start: number = match.index;
        const end: number = start + match[0].length;
        const room: string = match[1];

        return {
          room,
          start,
          end
        };
      }
    }

    // C. Any compact token followed by /###A.
    //
    // This also allows the right-side digits to contain spaces.
    {
      const re: RegExp =
        /\b[^\s/]{1,20}\s*\/\s*(\d)\s*(\d)\s*(\d)\s*([A-Z0-9])\b/g;

      const match: RegExpExecArray | null = re.exec(sp);

      if (match) {
        const start: number = match.index;
        const end: number = start + match[0].length;
        const room: string =
          `${match[1]}${match[2]}${match[3]}`;

        return {
          room,
          start,
          end
        };
      }
    }

    return null;
  };

  // ================= NAME EXTRACTION =================

  // Last name must follow the bed and precede the first comma.
  const findLastNameAfterBed = (
    afterSp: string,
    afterTi: string
  ): string => {
    const pre: string = afterSp.replace(
      /[~()]+/g,
      " "
    );

    const cleaned: string = pre.replace(
      /^[\s-]+/,
      ""
    );

    let match: RegExpMatchArray | null = cleaned.match(
      /^\s*([A-Za-z][A-Za-z' -]{1,40})\s*,/
    );

    if (match) {
      return match[1]
        .replace(/[()]/g, "")
        .trim();
    }

    const preTight: string = afterTi.replace(
      /[~()]+/g,
      ""
    );

    match = preTight.match(
      /^([A-Za-z][A-Za-z'-]{1,40}),/
    );

    if (match) {
      return match[1]
        .replace(/[()]/g, "")
        .trim();
    }

    return "";
  };

  // Extract the patient's first name for the patient de-dup key.
  //
  // For lines like:
  // Goode, Acute HM Chandler 12 No ... UNI Christopher Neutropenic
  //
  // this returns Christopher instead of Acute.
  const extractFirstNameAfterBed = (
    afterSp: string
  ): string => {
    const cleaned: string = afterSp
      .replace(/[~()]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    // Highest priority:
    // UNI <FirstName> Neutropenic
    const uniNeutropenicMatch:
      RegExpMatchArray | null = cleaned.match(
        /\bUNI\s+([A-Za-z][A-Za-z'-]{1,40})\s+Neutropenic\b/i
      );

    if (uniNeutropenicMatch) {
      return uniNeutropenicMatch[1].trim();
    }

    // Additional UNI fallback:
    // Allows UNI <FirstName> even if Neutropenic is truncated or missing.
    const uniNameMatch:
      RegExpMatchArray | null = cleaned.match(
        /\bUNI\s+([A-Za-z][A-Za-z'-]{1,40})\b/i
      );

    if (uniNameMatch) {
      const uniCandidate: string =
        uniNameMatch[1].trim();

      const invalidUniTokens = new Set<string>([
        "ACUTE",
        "PROGRESSIVE",
        "PROG",
        "PROGR",
        "PROGRESS",
        "ICU",
        "UNI",
        "YES",
        "NO",
        "CONTACT",
        "NEUTROPENIC"
      ]);

      if (
        !invalidUniTokens.has(
          uniCandidate.toUpperCase()
        )
      ) {
        return uniCandidate;
      }
    }

    // Standard format:
    // Last, First
    const normalNameMatch:
      RegExpMatchArray | null = cleaned.match(
        /^[A-Za-z][A-Za-z' -]{0,40}\s*,\s*([A-Za-z][A-Za-z'-]{0,40})\b/
      );

    if (!normalNameMatch) {
      return "";
    }

    const candidate: string =
      normalNameMatch[1].trim();

    // Clinical, service, and status tokens must not be used
    // as a patient's first name.
    const nonNameTokens = new Set<string>([
      "ACUTE",
      "PROGRESSIVE",
      "PROG",
      "PROGR",
      "PROGRESS",
      "ICU",
      "UNI",
      "HM",
      "SCR",
      "SGR",
      "GYO",
      "GYN",
      "SGO",
      "URO",
      "OMF",
      "HNS",
      "ENT",
      "VRE",
      "SGE",
      "SE",
      "TXP",
      "ORT",
      "SGT",
      "GEMS",
      "STICU",
      "SURGERY",
      "TRAUMA",
      "YES",
      "NO",
      "CONTACT",
      "NEUTROPENIC"
    ]);

    if (
      nonNameTokens.has(candidate.toUpperCase())
    ) {
      return "";
    }

    return candidate;
  };

  // ================= DIAGNOSIS =================

  const detectDiagnosis = (
    slice: string
  ): string => {
    if (
      /\bProg(r(ess(ive)?)?)?\b/i.test(slice)
    ) {
      return "Progressive";
    }

    if (
      /\bAcu(t(e)?)?\b/i.test(slice)
    ) {
      return "Acute";
    }

    if (
      /\bProgres/i.test(slice)
    ) {
      return "Progressive";
    }

    if (
      /\bAcute\b/i.test(slice)
    ) {
      return "Acute";
    }

    if (
      /\bICU\b/i.test(slice)
    ) {
      return "ICU";
    }

    return "";
  };

  // ================= SERVICE =================

  const serviceMap: Record<string, string> = {
    SCR: "SCR",
    SGR: "SGR",
    GYO: "GYO",
    GYN: "Gynecology",
    SGO: "SGO",
    URO: "URO",
    OMF: "OMF",
    HNS: "HNS",
    ENT: "ENT",
    VRE: "VRE",
    SGE: "SGE",
    SE: "SGE",
    TXP: "TXP",
    HM: "HM",
    ORT: "ORT",
    SGT: "SGT",
    GEMS: "GEMS",
    STICU: "SGT",
    SURGERY: "SGT",
    TRAUMA: "SGT"
  };

  const detectService = (
    slice: string
  ): string => {
    const tight: string = slice
      .replace(/\s+/g, "")
      .toUpperCase();

    const serviceKeys: string[] =
      Object.keys(serviceMap);

    for (const key of serviceKeys) {
      const upperKey: string = key.toUpperCase();

      const spacedPattern = new RegExp(
        `\\b${key}\\b`,
        "i"
      );

      if (
        spacedPattern.test(slice) ||
        tight.includes(upperKey)
      ) {
        return serviceMap[key];
      }
    }

    return "";
  };

  // ================= YES/NO =================

  const firstTwoYesNo = (
    slice: string
  ): [string, string] => {
    const output: string[] = [];

    const re: RegExp = /\b(Yes|No)\b/gi;
    let match: RegExpExecArray | null;

    while (
      (match = re.exec(slice)) !== null &&
      output.length < 2
    ) {
      output.push(match[1]);
    }

    // Handle joined tokens such as NoNo.
    if (output.length < 2) {
      const tight: string = slice
        .replace(/\s+/g, "")
        .toUpperCase();

      let i: number = 0;

      while (
        i < tight.length &&
        output.length < 2
      ) {
        if (tight.startsWith("YES", i)) {
          output.push("Yes");
          i += 3;
          continue;
        }

        if (tight.startsWith("NO", i)) {
          output.push("No");
          i += 2;
          continue;
        }

        i++;
      }
    }

    return [
      output[0] || "",
      output[1] || ""
    ];
  };

  // ================= SHEETS =================

  const sourceSheet =
    workbook.getWorksheet(SOURCE_SHEET_NAME);

  if (!sourceSheet) {
    console.log(
      `ERROR: Source sheet "${SOURCE_SHEET_NAME}" not found.`
    );
    return;
  }

  let destSheet =
    workbook.getWorksheet(DEST_SHEET_NAME);

  if (!destSheet) {
    destSheet =
      workbook.addWorksheet(DEST_SHEET_NAME);
  }

  // Clear the destination once.
  const destinationUsedRange =
    destSheet.getUsedRange();

  if (destinationUsedRange) {
    destinationUsedRange.clear(
      ExcelScript.ClearApplyTo.all
    );
  }

  // ================= SOURCE DATA =================

  const usedRange =
    sourceSheet.getUsedRange(true) ||
    sourceSheet.getUsedRange();

  if (!usedRange) {
    console.log(
      `ERROR: Source sheet "${SOURCE_SHEET_NAME}" has no used range.`
    );
    return;
  }

  const values: Cell[][] =
    usedRange.getValues() as Cell[][];

  // Build normalized spaced lines.
  const linesSp: string[] = values
    .map((row: Cell[]): string => {
      return toSpaced(row.join(" "));
    })
    .filter((line: string): boolean => {
      return (
        !/^\s*CH\s*$/i.test(line) &&
        !/^\s*CH\s*[~\-\|\(\)]*\s*$/i.test(line)
      );
    })
    .filter((line: string): boolean => {
      return line !== "";
    })
    .map((line: string): string => {
      return applyOcrTokenFixesSpaced(line);
    });

  // Build tight lines from the corrected spaced lines.
  const linesTi: string[] = linesSp
    .map((line: string): string => {
      return toTight(line);
    })
    .map((line: string): string => {
      return applyOcrTokenFixesTight(line);
    });

  if (DEBUG) {
    console.log(
      "Preview of the first 200 non-empty lines after OCR fixes:"
    );

    for (
      let i = 0;
      i < Math.min(linesSp.length, 200);
      i++
    ) {
      console.log(
        `${i + 1}: ${linesSp[i]}`
      );
    }
  }

  // ================= PARSE =================

  const collected: RowEx[] = [];

  let skipBoiler: number = 0;
  let skipNoBed: number = 0;
  let skipNoFloor: number = 0;
  let skipRange: number = 0;
  let skipNoName: number = 0;
  let skipNoisyCH: number = 0;

  for (
    let i = 0;
    i < linesSp.length;
    i++
  ) {
    const line1Spaced: string = linesSp[i];
    const line1Tight: string = linesTi[i];

    if (isBoilerplate(line1Spaced)) {
      if (DEBUG) {
        console.log(
          "[SKIP boilerplate]",
          i,
          ":",
          line1Spaced
        );
      }

      skipBoiler++;
      continue;
    }

    const line2Spaced: string =
      i + 1 < linesSp.length &&
      !isBoilerplate(linesSp[i + 1])
        ? linesSp[i + 1]
        : "";

    const line3Spaced: string =
      i + 2 < linesSp.length &&
      !isBoilerplate(linesSp[i + 2])
        ? linesSp[i + 2]
        : "";

    const line2Tight: string =
      i + 1 < linesTi.length
        ? linesTi[i + 1]
        : "";

    const line3Tight: string =
      i + 2 < linesTi.length
        ? linesTi[i + 2]
        : "";

    const windowSpaced: string =
      `${line1Spaced} ${line2Spaced} ${line3Spaced}`
        .replace(/\s+/g, " ")
        .trim();

    const windowTight: string =
      `${line1Tight}${line2Tight}${line3Tight}`;

    if (hasNoisyCH(windowSpaced)) {
      if (DEBUG) {
        console.log(
          "[SKIP noisy CH~~~]",
          i,
          ":",
          windowSpaced.slice(0, 180)
        );
      }

      skipNoisyCH++;
      continue;
    }

    // Require a bed in the three-line window.
    const bed: RoomHit | null =
      findBedWithBounds(
        windowSpaced,
        windowTight
      );

    if (!bed) {
      if (DEBUG) {
        console.log(
          "[SKIP no bed]",
          i,
          ":",
          windowSpaced.slice(0, 220)
        );
      }

      skipNoBed++;
      continue;
    }

    // Look around the bed for the floor header.
    const floorSliceStart: number =
      Math.max(
        0,
        bed.start - FLOOR_SLICE_BEFORE
      );

    const floorSliceEnd: number =
      Math.min(
        windowSpaced.length,
        bed.end + FLOOR_SLICE_AFTER
      );

    const floorSliceSpaced: string =
      windowSpaced.slice(
        floorSliceStart,
        floorSliceEnd
      );

    const floorSliceTight: string =
      floorSliceSpaced.replace(/\s+/g, "");

    const floorString: string =
      findFloorStrict(
        floorSliceSpaced,
        floorSliceTight
      );

    if (!floorString) {
      if (DEBUG) {
        console.log(
          "[SKIP no floor]",
          i,
          "around bed:",
          floorSliceSpaced.slice(0, 160)
        );
      }

      skipNoFloor++;
      continue;
    }

    const floorNumber: number =
      parseInt(floorString, 10);

    if (
      !(
        floorNumber >= FLOOR_MIN &&
        floorNumber <= FLOOR_MAX
      )
    ) {
      if (DEBUG) {
        console.log(
          "[SKIP floor out of range]",
          i,
          "floor:",
          floorNumber,
          "| window:",
          windowSpaced.slice(0, 160)
        );
      }

      skipRange++;
      continue;
    }

    // Read immediately after the bed.
    const afterEnd: number =
      Math.min(
        windowSpaced.length,
        bed.end + AFTER_SLICE_LEN
      );

    const afterSpaced: string =
      windowSpaced.slice(
        bed.end,
        afterEnd
      );

    const afterTight: string =
      afterSpaced.replace(/\s+/g, "");

    const lastName: string =
      findLastNameAfterBed(
        afterSpaced,
        afterTight
      );

    if (!lastName) {
      if (DEBUG) {
        console.log(
          "[SKIP no last name]",
          i,
          "after bed:",
          afterSpaced.slice(0, 140)
        );
      }

      skipNoName++;
      continue;
    }

    // This now finds Christopher in:
    // Goode, Acute ... UNI Christopher Neutropenic
    const firstName: string =
      extractFirstNameAfterBed(afterSpaced);

    // Build patient de-dup key.
    //
    // If the first name cannot be determined, use the room
    // rather than using a clinical token as the name.
    const patientKey: string =
      lastName.toUpperCase() +
      "|" +
      (
        firstName
          ? firstName.toUpperCase()
          : `RM${bed.room}`
      );

    const diagnosis: string =
      detectDiagnosis(afterSpaced);

    let service: string =
      detectService(afterSpaced);

    if (
      /^(YES|NO)$/i.test(service) ||
      service.toUpperCase() === "UNI"
    ) {
      service = "";
    }

    const yesNoValues: [string, string] =
      firstTwoYesNo(afterSpaced);

    const yn1: string = yesNoValues[0];
    const yn2: string = yesNoValues[1];

    if (DEBUG) {
      console.log(
        "[ADD]",
        {
          floorNumber,
          room: bed.room,
          lastName,
          firstName,
          diagnosis,
          service,
          yn1,
          yn2,
          patientKey
        }
      );
    }

    collected.push([
      String(floorNumber),
      bed.room,
      lastName,
      diagnosis,
      service,
      yn1,
      yn2,
      patientKey
    ]);
  }

  // ================= DEDUP #1: FLOOR + ROOM =================

  const bestByRoom =
    new Map<string, RowEx>();

  const score = (
    row: Row | RowEx
  ): number => {
    return (
      (row[2] ? 2 : 0) +
      (row[3] ? 1 : 0) +
      (row[4] ? 1 : 0) +
      (row[5] ? 0.5 : 0) +
      (row[6] ? 0.5 : 0)
    );
  };

  for (const row of collected) {
    const key: string =
      `${row[0]}|${row[1]}`;

    const previous: RowEx | undefined =
      bestByRoom.get(key);

    if (
      !previous ||
      score(row) > score(previous)
    ) {
      bestByRoom.set(key, row);
    }
  }

  const dedupedByRoom: RowEx[] =
    Array.from(bestByRoom.values());

  if (DEBUG) {
    console.log(
      "[DEDUP Room] from",
      collected.length,
      "to",
      dedupedByRoom.length
    );
  }

  // ================= DEDUP #2: PATIENT =================

  const bestByPatient =
    new Map<string, RowEx>();

  const frequencyByPatient =
    new Map<string, number>();

  for (const row of dedupedByRoom) {
    const key: string = row[7] || "";

    if (!key) {
      continue;
    }

    frequencyByPatient.set(
      key,
      (frequencyByPatient.get(key) || 0) + 1
    );
  }

  for (const row of dedupedByRoom) {
    const key: string = row[7] || "";

    if (!key) {
      continue;
    }

    const previous: RowEx | undefined =
      bestByPatient.get(key);

    if (
      !previous ||
      score(row) > score(previous)
    ) {
      bestByPatient.set(key, row);
    }
  }

  if (DEBUG) {
    const frequencyEntries:
      [string, number][] =
        Array.from(
          frequencyByPatient.entries()
        );

    for (const entry of frequencyEntries) {
      const key: string = entry[0];
      const occurrenceCount: number = entry[1];

      if (occurrenceCount > 1) {
        console.log(
          "[DEDUP Patient]",
          key,
          "occurrences:",
          occurrenceCount
        );
      }
    }
  }

  const patientDeduped: RowEx[] =
    Array.from(bestByPatient.values());

  const rowsForOutput: Row[] =
    patientDeduped.map(
      (row: RowEx): Row => {
        return [
          row[0],
          row[1],
          row[2],
          row[3],
          row[4],
          row[5],
          row[6]
        ];
      }
    );

  if (DEBUG) {
    console.log(
      "[DEDUP Patient] from",
      dedupedByRoom.length,
      "to",
      rowsForOutput.length
    );
  }

  // ================= SORT =================

  rowsForOutput.sort(
    (a: Row, b: Row): number => {
      const floorDifference: number =
        Number(a[0]) - Number(b[0]);

      if (floorDifference !== 0) {
        return floorDifference;
      }

      return Number(a[1]) - Number(b[1]);
    }
  );

  // ================= WRITE =================

  if (rowsForOutput.length > 0) {
    const outputRange =
      destSheet.getRangeByIndexes(
        0,
        0,
        rowsForOutput.length,
        7
      );

    outputRange.setValues(rowsForOutput);

    outputRange.getSort().apply([
      {
        key: 0,
        ascending: true
      },
      {
        key: 1,
        ascending: true
      }
    ]);
  } else {
    destSheet
      .getRange("A1")
      .setValue(
        "No rows emitted. Check logs."
      );
  }

  // ================= LOGS =================

  console.log(
    `Rows parsed: ${collected.length} -> after unique(Floor,Room): ${dedupedByRoom.length}`
  );

  console.log(
    `Then after unique(Patient): ${rowsForOutput.length}`
  );

  console.log(
    `Skipped boilerplate: ${skipBoiler}`
  );

  console.log(
    `Skipped noisy CH~~~ windows: ${skipNoisyCH}`
  );

  console.log(
    `Skipped (no bed ###/###X or duplicate pattern): ${skipNoBed}`
  );

  console.log(
    `Skipped (no strict PAV(A)+digits near bed): ${skipNoFloor}`
  );

  console.log(
    `Skipped (floor out of range ${FLOOR_MIN}-${FLOOR_MAX}): ${skipRange}`
  );

  console.log(
    `Skipped (no anchored last name after bed): ${skipNoName}`
  );

  // ================= ACTIVATE CHARGE SHEET =================

  const chargeSheet =
    workbook.getWorksheet("CHARGE_SHEET");

  if (chargeSheet) {
    chargeSheet.activate();
  } else {
    console.log(
      'WARNING: Worksheet "CHARGE_SHEET" was not found.'
    );
  }
}


SHAPE_TYPE_MAP = {
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "Embedded OLE Object",
    8: "Form Control",
    9: "Line",
    10: "Linked OLE Object",
    11: "Linked Picture",
    12: "OLE Control Object",
    13: "Picture",
    14: "Placeholder",
    15: "Text Effect",
    16: "Media",
    17: "Text Box",
    18: "Script Anchor",
    19: "Table",
    20: "Canvas",
    21: "Diagram",
    22: "Ink",
    23: "Ink Comment",
    24: "Smart Art",
    25: "Web Video",
    26: "Content App"
}

PLACEHOLDER_TYPE_MAP = {
    1: "Title",
    2: "Body",
    3: "CenterTitle",
    4: "SubTitle",
    5: "VerticalTitle",
    6: "VerticalBody",
    7: "Object",
    8: "Chart",
    9: "Table",
    10: "ClipArt",
    11: "OrgChart",
    12: "Media",
    13: "VerticalObject",
    14: "Picture",
    15: "Slide Number",
    16: "Header",
    17: "Footer",
    18: "Date",
    19: "VerticalTitle2",
    20: "VerticalBody2" 
}




# Source: https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype
# "Value": ["Name","Description"]
AUTOSHAPE_TYPE_MAP = {
  "1": ["msoShapeRectangle","Rectangle"],
  "2": ["msoShapeParallelogram","Parallelogram"],
  "3": ["msoShapeTrapezoid","Trapezoid"],
  "4": ["msoShapeDiamond","Diamond"],
  "5": ["msoShapeRoundedRectangle","Rounded rectangle"],
  "6": ["msoShapeOctagon","Octagon"],
  "7": ["msoShapeIsoscelesTriangle","Isosceles triangle"],
  "8": ["msoShapeRightTriangle","Right triangle"],
  "9": ["msoShapeOval","Oval"],
  "10": ["msoShapeHexagon","Hexagon"],
  "11": ["msoShapeCross","Cross"],
  "12": ["msoShapeRegularPentagon","Pentagon"],
  "13": ["msoShapeCan","Can"],
  "14": ["msoShapeCube","Cube"],
  "15": ["msoShapeBevel","Bevel"],
  "16": ["msoShapeFoldedCorner","Folded corner"],
  "17": ["msoShapeSmileyFace","Smiley face"],
  "18": ["msoShapeDonut","Donut"],
  "19": ["msoShapeNoSymbol","'No' symbol"],
  "20": ["msoShapeBlockArc","Block arc"],
  "21": ["msoShapeHeart","Heart"],
  "22": ["msoShapeLightningBolt","Lightning bolt"],
  "23": ["msoShapeSun","Sun"],
  "24": ["msoShapeMoon","Moon"],
  "25": ["msoShapeArc","Arc"],
  "26": ["msoShapeDoubleBracket","Double bracket"],
  "27": ["msoShapeDoubleBrace","Double brace"],
  "28": ["msoShapePlaque","Plaque"],
  "29": ["msoShapeLeftBracket","Left bracket"],
  "30": ["msoShapeRightBracket","Right bracket"],
  "31": ["msoShapeLeftBrace","Left brace"],
  "32": ["msoShapeRightBrace","Right brace"],
  "33": ["msoShapeRightArrow","Block arrow that points right"],
  "34": ["msoShapeLeftArrow","Block arrow that points left"],
  "35": ["msoShapeUpArrow","Block arrow that points up"],
  "36": ["msoShapeDownArrow","Block arrow that points down"],
  "37": ["msoShapeLeftRightArrow","Block arrow with arrowheads that point both left and right"],
  "38": ["msoShapeUpDownArrow","Block arrow that points up and down"],
  "39": ["msoShapeQuadArrow","Block arrows that point up, down, left, and right"],
  "40": ["msoShapeLeftRightUpArrow","Block arrow with arrowheads that point left, right, and up"],
  "41": ["msoShapeBentArrow","Block arrow that follows a curved 90-degree angle."],
  "42": ["msoShapeUTurnArrow","Block arrow forming a U shape"],
  "43": ["msoShapeLeftUpArrow","Block arrow with arrowheads that point left and up"],
  "44": ["msoShapeBentUpArrow","Block arrow that follows a sharp 90-degree angle. Points up by default."],
  "45": ["msoShapeCurvedRightArrow","Block arrow that curves right"],
  "46": ["msoShapeCurvedLeftArrow","Block arrow that curves left"],
  "47": ["msoShapeCurvedUpArrow","Block arrow that curves up"],
  "48": ["msoShapeCurvedDownArrow","Block arrow that curves down"],
  "49": ["msoShapeStripedRightArrow","Block arrow that points right with stripes at the tail"],
  "50": ["msoShapeNotchedRightArrow","Notched block arrow that points right"],
  "51": ["msoShapePentagon","Pentagon"],
  "52": ["msoShapeChevron","Chevron"],
  "53": ["msoShapeRightArrowCallout","Callout with arrow that points right"],
  "54": ["msoShapeLeftArrowCallout","Callout with arrow that points left"],
  "55": ["msoShapeUpArrowCallout","Callout with arrow that points up"],
  "56": ["msoShapeDownArrowCallout","Callout with arrow that points down"],
  "57": ["msoShapeLeftRightArrowCallout","Callout with arrowheads that point both left and right"],
  "58": ["msoShapeUpDownArrowCallout","Callout with arrows that point up and down"],
  "59": ["msoShapeQuadArrowCallout","Callout with arrows that point up, down, left, and right"],
  "60": ["msoShapeCircularArrow","Block arrow that follows a curved 180-degree angle"],
  "61": ["msoShapeFlowchartProcess","Process flowchart symbol"],
  "62": ["msoShapeFlowchartAlternateProcess","Alternate process flowchart symbol"],
  "63": ["msoShapeFlowchartDecision","Decision flowchart symbol"],
  "64": ["msoShapeFlowchartData","Data flowchart symbol"],
  "65": ["msoShapeFlowchartPredefinedProcess","Predefined process flowchart symbol"],
  "66": ["msoShapeFlowchartInternalStorage","Internal storage flowchart symbol"],
  "67": ["msoShapeFlowchartDocument","Document flowchart symbol"],
  "68": ["msoShapeFlowchartMultidocument","Multi-document flowchart symbol"],
  "69": ["msoShapeFlowchartTerminator","Terminator flowchart symbol"],
  "70": ["msoShapeFlowchartPreparation","Preparation flowchart symbol"],
  "71": ["msoShapeFlowchartManualInput","Manual input flowchart symbol"],
  "72": ["msoShapeFlowchartManualOperation","Manual operation flowchart symbol"],
  "73": ["msoShapeFlowchartConnector","Connector flowchart symbol"],
  "74": ["msoShapeFlowchartOffpageConnector","Off-page connector flowchart symbol"],
  "75": ["msoShapeFlowchartCard","Card flowchart symbol"],
  "76": ["msoShapeFlowchartPunchedTape","Punched tape flowchart symbol"],
  "77": ["msoShapeFlowchartSummingJunction","Summing junction flowchart symbol"],
  "78": ["msoShapeFlowchartOr","'Or' flowchart symbol"],
  "79": ["msoShapeFlowchartCollate","Collate flowchart symbol"],
  "80": ["msoShapeFlowchartSort","Sort flowchart symbol"],
  "81": ["msoShapeFlowchartExtract","Extract flowchart symbol"],
  "82": ["msoShapeFlowchartMerge","Merge flowchart symbol"],
  "83": ["msoShapeFlowchartStoredData","Stored data flowchart symbol"],
  "84": ["msoShapeFlowchartDelay","Delay flowchart symbol"],
  "85": ["msoShapeFlowchartSequentialAccessStorage","Sequential access storage flowchart symbol"],
  "86": ["msoShapeFlowchartMagneticDisk","Magnetic disk flowchart symbol"],
  "87": ["msoShapeFlowchartDirectAccessStorage","Direct access storage flowchart symbol"],
  "88": ["msoShapeFlowchartDisplay","Display flowchart symbol"],
  "89": ["msoShapeExplosion1","Explosion"],
  "90": ["msoShapeExplosion2","Explosion"],
  "91": ["msoShape4pointStar","4-point star"],
  "92": ["msoShape5pointStar","5-point star"],
  "93": ["msoShape8pointStar","8-point star"],
  "94": ["msoShape16pointStar","16-point star"],
  "95": ["msoShape24pointStar","24-point star"],
  "96": ["msoShape32pointStar","32-point star"],
  "97": ["msoShapeUpRibbon","Ribbon banner with center area above ribbon ends"],
  "98": ["msoShapeDownRibbon","Ribbon banner with center area below ribbon ends"],
  "99": ["msoShapeCurvedUpRibbon","Ribbon banner that curves up"],
  "100": ["msoShapeCurvedDownRibbon","Ribbon banner that curves down"],
  "101": ["msoShapeVerticalScroll","Vertical scroll"],
  "102": ["msoShapeHorizontalScroll","Horizontal scroll"],
  "103": ["msoShapeWave","Wave"],
  "104": ["msoShapeDoubleWave","Double wave"],
  "105": ["msoShapeRectangularCallout","Rectangular callout"],
  "106": ["msoShapeRoundedRectangularCallout","Rounded rectangle-shaped callout"],
  "107": ["msoShapeOvalCallout","Oval-shaped callout"],
  "108": ["msoShapeCloudCallout","Cloud callout"],
  "109": ["msoShapeLineCallout1","Callout with border and horizontal callout line"],
  "110": ["msoShapeLineCallout2","Callout with diagonal straight line"],
  "111": ["msoShapeLineCallout3","Callout with angled line"],
  "112": ["msoShapeLineCallout4","Callout with callout line segments forming a U-shape"],
  "113": ["msoShapeLineCallout1AccentBar","Callout with horizontal accent bar"],
  "114": ["msoShapeLineCallout2AccentBar","Callout with diagonal callout line and accent bar"],
  "115": ["msoShapeLineCallout3AccentBar","Callout with angled callout line and accent bar"],
  "116": ["msoShapeLineCallout4AccentBar","Callout with accent bar and callout line segments forming a U-shape"],
  "117": ["msoShapeLineCallout1NoBorder","Callout with horizontal line"],
  "118": ["msoShapeLineCallout2NoBorder","Callout with no border and diagonal callout line"],
  "119": ["msoShapeLineCallout3NoBorder","Callout with no border and angled callout line"],
  "120": ["msoShapeLineCallout4NoBorder","Callout with no border and callout line segments forming a U-shape"],
  "121": ["msoShapeLineCallout1BorderandAccentBar","Callout with border and horizontal accent bar"],
  "122": ["msoShapeLineCallout2BorderandAccentBar","Callout with border, diagonal straight line, and accent bar"],
  "123": ["msoShapeLineCallout3BorderandAccentBar","Callout with border, angled callout line, and accent bar"],
  "124": ["msoShapeLineCallout4BorderandAccentBar","Callout with border, accent bar, and callout line segments forming a U-shape"],
  "125": ["msoShapeActionButtonCustom","Button with no default picture or text. Supports mouse-click and mouse-over actions."],
  "126": ["msoShapeActionButtonHome","Home button. Supports mouse-click and mouse-over actions."],
  "127": ["msoShapeActionButtonHelp","Help button. Supports mouse-click and mouse-over actions."],
  "128": ["msoShapeActionButtonInformation","Information button. Supports mouse-click and mouse-over actions."],
  "129": ["msoShapeActionButtonBackorPrevious","Back or Previous button. Supports mouse-click and mouse-over actions."],
  "130": ["msoShapeActionButtonForwardorNext","Forward or Next button. Supports mouse-click and mouse-over actions."],
  "131": ["msoShapeActionButtonBeginning","Beginning button. Supports mouse-click and mouse-over actions."],
  "132": ["msoShapeActionButtonEnd","End button. Supports mouse-click and mouse-over actions."],
  "133": ["msoShapeActionButtonReturn","Return button. Supports mouse-click and mouse-over actions."],
  "134": ["msoShapeActionButtonDocument","Document button. Supports mouse-click and mouse-over actions."],
  "135": ["msoShapeActionButtonSound","Sound button. Supports mouse-click and mouse-over actions."],
  "136": ["msoShapeActionButtonMovie","Movie button. Supports mouse-click and mouse-over actions."],
  "137": ["msoShapeBalloon","Balloon"],
  "138": ["msoShapeNotPrimitive","Not supported"],
  "139": ["msoShapeFlowchartOfflineStorage","Offline storage flowchart symbol"],
  "140": ["msoShapeLeftRightRibbon","Ribbon with an arrow at both ends"],
  "141": ["msoShapeDiagonalStripe","Rectangle with two triangles-shapes removed; a diagonal stripe"],
  "142": ["msoShapePie","Circle ('pie') with a portion missing"],
  "143": ["msoShapeNonIsoscelesTrapezoid","Trapezoid with asymmetrical non-parallel sides"],
  "144": ["msoShapeDecagon","Decagon"],
  "145": ["msoShapeHeptagon","Heptagon"],
  "146": ["msoShapeDodecagon","Dodecagon"],
  "147": ["msoShape6pointStar","6-point star"],
  "148": ["msoShape7pointStar","7-point star"],
  "149": ["msoShape10pointStar","10-point star"],
  "150": ["msoShape12pointStar","12-point star"],
  "151": ["msoShapeRound1Rectangle","Rectangle with one rounded corner"],
  "152": ["msoShapeRound2SameRectangle","Rectangle with two-rounded corners that share a side"],
  "154": ["msoShapeSnipRoundRectangle","Rectangle with one snipped corner and one rounded corner"],
  "155": ["msoShapeSnip1Rectangle","Rectangle with one snipped corner"],
  "156": ["msoShapeSnip2SameRectangle","Rectangle with two snipped corners that share a side"],
  "157": ["msoShapeSnip2DiagRectangle","Rectangle with two snipped corners, diagonally-opposed"],
  "158": ["msoShapeFrame","Rectangular picture frame"],
  "159": ["msoShapeHalfFrame","Half of a rectangular picture frame"],
  "160": ["msoShapeTear","Water droplet"],
  "161": ["msoShapeChord","Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord"],
  "162": ["msoShapeCorner","Rectangle with rectangular-shaped hole."],
  "163": ["msoShapeMathPlus","Addition symbol +"],
  "164": ["msoShapeMathMinus","Subtraction symbol -"],
  "165": ["msoShapeMathMultiply","Multiplication symbol x"],
  "166": ["msoShapeMathDivide","Division symbol ÷"],
  "167": ["msoShapeMathEqual","Equivalence symbol ="],
  "168": ["msoShapeMathNotEqual","Non-equivalence symbol ≠"],
  "169": ["msoShapeCornerTabs","Four right triangles aligning along a rectangular path; four 'snipped' corners."],
  "170": ["msoShapeSquareTabs","Four small squares that define a rectangular shape"],
  "171": ["msoShapePlaqueTabs","Four quarter-circles defining a rectangular shape"],
  "172": ["msoShapeGear6","Gear with six teeth"],
  "173": ["msoShapeGear9","Gear with nine teeth"],
  "174": ["msoShapeFunnel","Funnel"],
  "175": ["msoShapePieWedge","Quarter of a circular shape"],
  "176": ["msoShapeLeftCircularArrow","Circular arrow pointing counter-clockwise"],
  "177": ["msoShapeLeftRightCircularArrow","Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends"],
  "178": ["msoShapeSwooshArrow","Curved arrow"],
  "179": ["msoShapeCloud","Cloud shape"],
  "180": ["msoShapeChartX","Square divided into four parts along diagonal lines"],
  "181": ["msoShapeChartStar","Square divided into six parts along vertical and diagonal lines"],
  "182": ["msoShapeChartPlus","Square divided vertically and horizontally into four quarters"],
  "183": ["msoShapeLineInverse","Line inverse"]
}

# AUTOSHAPE_TYPE_MAP

# Bullet Point styles

# ID: [Name, Example]
BULLET_STYLE_MAP = {
    0: ["Mixed", ""],
    1: ["AlphaLCPeriod", "a."],
    2: ["AlphaUCPeriod", "A."],
    3: ["ArabicParenRight", "1)"],
    4: ["ArabicPeriod", "1."],
    5: ["RomanLCParenBoth", "(i)"],
    6: ["RomanLCPeriod", "i."],
    7: ["RomanUCPeriod", "I."],
    8: ["AlphaLCParenBoth", "(a)"],
    9: ["AlphaLCParenRight", "a)"],
    10: ["AlphaUCParenBoth", "(A)"],
    11: ["AlphaUCParenRight", "A)"],
    12: ["ArabicParenBoth", "(1)"],
    13: ["ArabicPlain", "1"],
    14: ["RomanLCParenRight", "i)"],
    15: ["RomanUCParenBoth", "(I)"],
    16: ["RomanUCParenRight", "I)"],
}
BULLET_CHAR_MAP = {
    8226: ["Black Round Dot", "•"],
    8211: ["En Dash Bar", "–"],
    10004: ["Check Mark", "✔"],
    9632: ["Black Square", "■"],
    9675: ["Hollow Circle", "○"],
    10146: ["Right Arrow", "➢"],
    61623: ["Diamond", "◆"], # Wingdings mapped character
}


#Chart
CHART_TYPES = {
    "column": 51,          # xlColumnClustered
    "stacked_column": 52,  # xlColumnStacked
    "line": 4,             # xlLine
    "line_markers": 65,    # xlLineMarkers
    "pie": 5,              # xlPie
    "bar": 57,             # xlBarClustered
    "stacked_bar": 58,     # xlBarStacked
    "area": 1,             # xlArea
    "scatter": -4169       # xlXYScatter
}

# --- Legend Position (XlLegendPosition) ---
LEGEND_POS = {
    "top": -4160,          # xlLegendPositionTop
    "bottom": -4107,       # xlLegendPositionBottom
    "left": -4131,         # xlLegendPositionLeft
    "right": -4152,        # xlLegendPositionRight
    "corner": 2,           # xlLegendPositionCorner
    "none": 0              # custom: hide legend
}


# --- PpEntryEffect (slide transition entry effect) ---
# Source: verified Microsoft values from bugs-and-fixes (many online refs are wrong).
# Reverse map used by parse_slide_transition. Unknown raw values are surfaced as ints.
PP_ENTRY_EFFECT_NAME = {
    0: "none",
    257: "cut",
    1025: "checkerboard",
    1284: "cover",
    1537: "dissolve",
    1793: "fade",
    2052: "uncover",
    2305: "random_bars",
    2819: "wipe",
    3585: "split",
    3849: "fade_smoothly",
    3852: "push",
    3854: "morph",  # PowerPoint 2016+
}


# --- MsoAnimTriggerType (animation trigger) ---
# Source: https://learn.microsoft.com/en-us/office/vba/api/office.msoanimtriggertype
MSO_ANIM_TRIGGER_NAME = {
    1: "OnClick",
    2: "WithPrevious",
    3: "AfterPrevious",
}


# --- MsoAnimEffect (shape animation effect type) ---
# Verified subset (matches the effect_map in tools.add_animation). The full
# MsoAnimEffect enum has 100+ values with overlapping semantics across PPT
# versions; only the verified-correct mappings live here. Unknown raw values
# are surfaced as 'raw_<N>' by the parser for forensics.
MSO_ANIM_EFFECT_NAME = {
    0:  "Custom",
    1:  "Appear",
    10: "Fade",
    22: "Fly",
    88: "Zoom",
}

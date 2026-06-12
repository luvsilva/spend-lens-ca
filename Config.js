// ============================================================
// Config.gs — Constantes globais + Menu
// ============================================================
const CONFIG = {
  TARGET_SHEET:  "Expenses_Form",
  BUDGET_SHEET:  "Budgets",
  RULES_SHEET:   "CategoryRules",
  COLS: { TIMESTAMP:1, DATE:2, DESCRIPTION:3, CATEGORY:4, AMOUNT:5, BANK:6 },
  UNCATEGORIZED_COLOR: "#FFF9C4",
  INCOME_CATEGORIES: ["Transfer", "Plan it"],

  CATEGORIES: [
    // ── Groceries ──────────────────────────────────────────────────────────
    { match: ["COSTCO", "COST-CO"],                                     category: "Groceries" },
    { match: ["BUY LOW", "BUY-LOW"],                                    category: "Groceries" },
    { match: ["SAVE ON FOODS", "SAVE-ON-FOODS", "SAVEON"],              category: "Groceries" },
    { match: ["REAL CDN SUPER", "REAL CANADIAN"],                       category: "Groceries" },
    { match: ["LANGLEY FARM", "FARM MARKET"],                           category: "Groceries" },
    { match: ["T&T"],                                                    category: "Groceries" },
    { match: ["DOLLARAMA"],                                              category: "Groceries" },
    { match: ["WALMART", "WAL-MART"],                                   category: "Groceries" },
    { match: ["SHOPPERS DRUG MART", "SHOPPERS DRUG"],                   category: "Groceries" },
    { match: ["COBS BREAD"],                                            category: "Groceries" },
    { match: ["TOMMY'S MARKET"],                                        category: "Groceries" },
    { match: ["LONDON DRUGS"],                                          category: "Groceries" },
    // ── Restaurant ─────────────────────────────────────────────────────────
    { match: ["SUSHI ON"],                                              category: "Restaurant" },
    { match: ["QUESADA"],                                               category: "Restaurant" },
    { match: ["ROMER'S"],                                               category: "Restaurant" },
    { match: ["PINT PUBLIC HOUSE", "THE PINT"],                         category: "Restaurant" },
    { match: ["LULU ISLAND"],                                           category: "Restaurant" },
    { match: ["GRAND VILLA"],                                           category: "Restaurant" },
    { match: ["THE PATCH BREWERY"],                                     category: "Restaurant" },
    { match: ["AL PORTO"],                                              category: "Restaurant" },
    { match: ["WATASHI SUSHI"],                                         category: "Restaurant" },
    { match: ["LOCHDALE VILLAGE"],                                      category: "Restaurant" },
    { match: ["CACTUS CLUB"],                                           category: "Restaurant" },
    // ── Fast Food ──────────────────────────────────────────────────────────
    { match: ["MCDONALD", "MCDONALDS"],                                 category: "Fast Food" },
    { match: ["TIM HORTON"],                                            category: "Fast Food" },
    { match: ["FRESHSLICE"],                                            category: "Fast Food" },
    { match: ["BURGER KING"],                                           category: "Fast Food" },
    { match: ["A&W"],                                                   category: "Fast Food" },
    { match: ["SUBWAY"],                                                category: "Fast Food" },
    // ── Coffee Shop ────────────────────────────────────────────────────────
    { match: ["STARBUCKS"],                                             category: "Coffee Shop" },
    { match: ["BLENZ"],                                                 category: "Coffee Shop" },
    { match: ["WAVES COFFEE", "WAVES COF"],                             category: "Coffee Shop" },
    { match: ["TREES ORGANIC"],                                         category: "Coffee Shop" },
    { match: ["JJ BEAN"],                                               category: "Coffee Shop" },
    { match: ["LA FORET CAFE"],                                         category: "Coffee Shop" },
    { match: ["TAKE 5 CAFE"],                                           category: "Coffee Shop" },
    { match: ["EUREST-BALLARD", "EUREST BALLARD"],                      category: "Coffee Shop" },
    // ── GAS ────────────────────────────────────────────────────────────────
    { match: ["CHEVRON", "CHV4", "CHV5"],                               category: "GAS" },
    { match: ["SHELL"],                                                  category: "GAS" },
    { match: ["ESSO"],                                                   category: "GAS" },
    { match: ["PETRO"],                                                  category: "GAS" },
    { match: ["FAS GAS"],                                                category: "GAS" },
    { match: ["GILLEY ON THE"],                                          category: "GAS" },
    { match: ["CANADA WAY CH"],                                          category: "GAS" },
    // ── Streaming ──────────────────────────────────────────────────────────
    { match: ["APPLE.COM/BILL", "APPLECOMBILL"],                        category: "Streamming" },
    { match: ["NETFLIX"],                                               category: "Streamming" },
    { match: ["SPOTIFY"],                                               category: "Streamming" },
    { match: ["AMAZON.CA PRIME", "AMAZON PRIME", "PRIME MEMBER"],       category: "Streamming" },
    { match: ["GOOGLE ONE"],                                            category: "Streamming" },
    { match: ["EPIDEMIC SOUND"],                                        category: "Streamming" },
    { match: ["DISNEY PLUS", "DISNEY+"],                                category: "Streamming" },
    { match: ["UPPBEAT"],                                               category: "Streamming" },
    { match: ["PLAYSTATIONNETW"],                                       category: "Streamming" },
    { match: ["PIKA.ART"],                                              category: "Streamming" },
    // ── GYM ────────────────────────────────────────────────────────────────
    { match: ["CLUB16", "CLUB 16", "ABC*30922"],                        category: "GYM" },
    { match: ["BURNABY PRCS ECC"],                                      category: "GYM" },
    // ── Supplements ────────────────────────────────────────────────────────
    { match: ["BODY ENERGY CLUB"],                                      category: "Supplements" },
    { match: ["POPEYE'S SUPPLEMENTS"],                                  category: "Supplements" },
    { match: ["VIDA NOVA HEALTH"],                                      category: "Supplements" },
    // ── Mobile / Internet ──────────────────────────────────────────────────
    { match: ["FIDO"],                                                   category: "Mobile" },
    { match: ["TELUS"],                                                  category: "Internet" },
    // ── BCHydro ────────────────────────────────────────────────────────────
    { match: ["BC HYDRO", "B.C. HYDRO", "HYDRO & POWER AUTHORITY"],     category: "BCHydro" },
    // ── Car Payment ────────────────────────────────────────────────────────
    { match: ["SCOTIA BANK - SPL LOAN", "SPL LOAN PAYMENT"],            category: "Car Payment" },
    // ── Appartament Rent ───────────────────────────────────────────────────
    { match: ["ONE WEST PROP"],                                          category: "Appartament Rent" },
    { match: ["CHEQUE"],                                                 category: "Appartament Rent" },
    // ── Fees ───────────────────────────────────────────────────────────────
    { match: ["SERVICE CHARGE", "MONTHLY FEE", "MEMBERSHIP FEE"],       category: "Fees" },
    { match: ["OVERDRAFT", "OVER LIMIT FEE", "PAY PER USE OVERDRAFT"],  category: "Banking Fees" },
    // ── ICBC / Car Insurance ───────────────────────────────────────────────
    { match: ["INSURANCE CORPORATION OF BC", "IC_8D.QJQ"],              category: "ICBC" },
    { match: ["ALL WRITE INSURANCE"],                                   category: "Car Insurance" },
    { match: ["MAXXAM INSURANCE"],                                      category: "Car Insurance" },
    // ── Parking ────────────────────────────────────────────────────────────
    { match: ["PAYBYPHONE", "PAY BY PHONE"],                            category: "Parking" },
    { match: ["HONK - CAN"],                                            category: "Parking" },
    { match: ["IMPARK"],                                                category: "Parking" },
    { match: ["ADV PARKING"],                                           category: "Parking" },
    { match: ["INDIGO PARK"],                                           category: "Parking" },
    // ── Transportation ─────────────────────────────────────────────────────
    { match: ["COMPASS VENDING", "COMPASS WEB", "COMPASS AUTOLOA"],     category: "Public Transportation" },
    { match: ["TRANSLINK"],                                             category: "Public Transportation" },
    { match: ["UBER"],                                                  category: "Private Transportation" },
    // ── Car Maintenance ────────────────────────────────────────────────────
    { match: ["SHINE AUTO WASH"],                                       category: "Car Maintenance" },
    { match: ["CANADIAN TIRE"],                                         category: "Car Maintenance" },
    { match: ["BRAUTO AUTOMOTIVE"],                                     category: "Car Maintenance" },
    { match: ["SQ BRAZILIAN AUTO", "BRAZILIAN AUTO"],                   category: "Car Maintenance" },
    // ── Clothes ────────────────────────────────────────────────────────────
    { match: ["H&M"],                                                   category: "Clothes" },
    { match: ["ZARA"],                                                  category: "Clothes" },
    { match: ["OLD NAVY"],                                              category: "Clothes" },
    { match: ["URBAN BEHAVIOR"],                                        category: "Clothes" },
    { match: ["LULULEMON"],                                             category: "Clothes" },
    { match: ["WINNERS", "HOMESENSE"],                                  category: "Clothes" },
    { match: ["DECATHLON"],                                             category: "Clothes" },
    { match: ["SUNGLASS HUT"],                                          category: "Clothes" },
    { match: ["UNIQLO"],                                                category: "Clothes" },
    { match: ["NIKE"],                                                  category: "Clothes" },
    { match: ["UNDER ARMOUR"],                                          category: "Clothes" },
    { match: ["SHEIN"],                                                 category: "Clothes" },
    { match: ["FOOT LOCKER"],                                           category: "Clothes" },
    // ── SkinCare / Health ──────────────────────────────────────────────────
    { match: ["SEPHORA"],                                               category: "SkinCare" },
    { match: ["PHARMASAVE"],                                            category: "Pharmacy" },
    { match: ["LIFELABS"],                                              category: "Health" },
    { match: ["CONFIDENCE DENTAL"],                                     category: "Dental" },
    // ── House / Dog ────────────────────────────────────────────────────────
    { match: ["IKEA"],                                                  category: "House Maintanence" },
    { match: ["PETSMART"],                                              category: "Dog Food" },
    { match: ["PET VALU", "BOSLEY"],                                    category: "Dog Food" },
    { match: ["KING GEORGE VET", "VETERINARY", "VETERINARI"],           category: "Dog Food" },
    // ── Leisure / Vacation ─────────────────────────────────────────────────
    { match: ["BOOKING.COM"],                                           category: "Vacation" },
    { match: ["AIRBNB"],                                                category: "Vacation" },
    { match: ["PARKS/PARCS CANADA"],                                    category: "Leisure" },
    { match: ["CINEPLEX"],                                              category: "Leisure" },
    { match: ["GUILTY & COMPANY"],                                      category: "Night Club" },
    // ── BarberShop ─────────────────────────────────────────────────────────
    { match: ["ATHENS HAIR", "HAIR STYLING", "HAIR CUT"],               category: "BarberShop" },
    { match: ["ROBERT FELIX HAIR"],                                     category: "BarberShop" },
    // ── Finance / Misc ─────────────────────────────────────────────────────
    { match: ["WEALTHSIMPLE"],                                          category: "Tax Return" },
    { match: ["SQUARE ONE INSURANCE"],                                  category: "Rent Insurance" },
    { match: ["BOARD OF EDUCATION"],                                    category: "Courses" },
    { match: ["DRIVER SERVICES CENTRE"],                                category: "Driver School" },
    { match: ["BC LIQUOR"],                                             category: "BCLiquor" },
    // Amazon genérico → SEM categoria (classificação manual)
    // Apenas Amazon Prime é classificado (acima em Streamming)
    // { match: ["AMAZON.CA", "AMZN"], category: "..." } ← REMOVIDO
    // ── Transfers (excluídos da análise) ───────────────────────────────────
    { match: ["INTERAC E-TRANSFER", "E-TRANSFER", "INTERAC ETRANSFER"], category: "Transfer" },
    // ── ATM / Cash Withdrawal ──────────────────────────────────────
    { match: ["ATM WITHDRAWAL", "AUTOMATED BANKING MACHINE", "ABM WITHDRAWAL"], category: "Cash Withdrawal" },
    { match: ["ATM WITH"],                                                        category: "Cash Withdrawal" },
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 Budget Importer")
    .addItem("📂 Import Bank File",               "openReviewImportModalV5")
    .addSeparator()
    .addItem("🔄 Re-categorize uncategorized",    "recategorizeUncategorized")
    .addSeparator()
    .addItem("📋 Category Rules",                  "toggleCategoryRulesSheet")
    .addItem("🚫 Import Filters",                  "toggleImportFiltersSheet")
    .addItem("💵 Budget",                          "toggleBudgetSheet")
    .addItem("📖 User Manual",                     "toggleManualSheet")
    .addSeparator()
    .addItem("📤 Export Clean Template",          "exportCleanTemplate")
    .addToUi();
}
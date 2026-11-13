# Decision log (newest last)

- Flat $500/month for every member, no founding rate, 12-month commitment, cap 100.
- Onboarding (interview, testing, custom routine, 3-month retest) included in every membership.
- Strength sessions and lessons both $150 per session; independent coaches keep 80%.
- 24/7 keycard or app access restored, no front desk. Gym is 3 digital strength stations plus a functional training tower and non-barbell implements (liability).
- 3,000 sq ft used everywhere (model, deck, kit). Normatec/recovery $3K. Insurance $2,100 placeholder mentioned by Jeremy; model currently carries $800/month (see OPEN_ITEMS).
- Wyatt removed from the GM and pro roles. The PGA professional is unnamed. GM is a separate $45K hire covering operations, social media and member experience.
- Deck wording: avoid "athlete"; solution headline "Everything you need to prepare the swing and the body that powers it for optimal performance". Problem slide has four cards: scattered, siloed, guesswork, someone else's schedule.
- Added a "what sets us apart" slide, a price-comparison slide (about $487 separately vs $500) and a founder story slide (slide 3).
- Market slide competitors: Back Nine, Albatross, Choo Choo, one TPI-certified PT (session or package only). Private clubs are partners and a member source, not competitors.
- Advising partner removed from the team slide; it stays on the ask slide.
- Nov 10 audit: strength-coach follow-up pay ($22.50 per member per month) now applies to all members (it was skipped for the first 25). Break-even about 38 members, contribution $527, Base Year 3 profit $215K, capital $385K Base and $433K Conservative.

## Nov 10, 2026
- Inputs: Tonals quoted $11K each with 5-year insurance/service warranty paid upfront; tower, weights and storage $15K; treadmill $4K; assault bike $3K (gym $55K). Insurance estimate $185/month (liability, 24/7 unstaffed, no building; formal quote pending). Rent still the $12/sq ft placeholder.
- Dropped the $1,500 velocity device and ski erg (not in the new list).
- Removed staged funding from the ask (slide 13) and risks (slide 14). The $500K is one raise; unused capital still returned in Year 2.
- Slide 11 combines ramp operating losses and safety cushion into one row ("Operating losses until break-even, plus safety cushion": $69K Base, $108K Conservative).

## Nov 11, 2026
- Rent planned at $16/sq ft/year all in ($4,000/month), a more conservative placeholder; no flex industrial quote yet.
- Unused capital is no longer returned in Year 2. It stays in the account as a buffer, drawn as needed.
- Investor sale/exit and buyout scenarios removed (exit multiple, "with exit" returns, IRR, Ownership scenarios sheet). Returns are cash from operations only. Revisit if/when needed.
- Staged funding (tranches) removed from the model and deck. The old tranches were line-item groupings, not a spend timeline, so they did not guarantee cash before spending. Funds are drawn as needed.
- Sources are the investor raise only (no owner cash, no founding-member cash).
- Explanatory text stripped from the workbook (notes columns, prose rows, Guide and Changes and flags sheets). Built by model/simplify.py, which runs last.

## Nov 11, 2026 (later)
- Model is dues only: member take-up of coaching and lessons set to 0. Contractor take-home rows removed. Training-only packages ($175 PGA lesson, $199 performance/athletic therapy, $150 personal training/strength, per hour) sit on Assumptions rows 209-212, not in the projections.
- Owner takes no distributions until the investor is repaid: investor share until payback = 100% (Assumptions B57). Straight equity and 1.5x structures removed because they do not repay within 5 years.
- Open: follow-up coaching is costed at $30/hr (Assumptions B190), a placeholder. At $60/hr Base Y3 profit drops about $21K; at $120/hr about $54K more.

## Nov 11, 2026 (cap)
- Member cap raised from 100 to 133: 90% of the 148.7 members the five stations carry at the 85% peak ceiling (Station capacity tab; buffer input B17). Upside now reaches the cap by about month 34. Space is about 22.5 sq ft per member at the cap.

## Nov 11, 2026 (TrackMan, commitment, churn)
- TrackMan expansion logic removed from the model (five hitting stations only). This also removes a $250/month "added bay" software cost that was being charged by a leftover switch.
- All memberships are 3-month commitments, marketed as "give us 3 months to transform your game". Churn planned at 5% a month in the Base case (Conservative 6%, Upside 4%, keeping the 1-point spread); working to bring it lower. The end-of-quarter retest is the retention moment.
- Term discounts: 5% off the monthly price for a 6-month term, 10% off for a 12-month term. Inputs on Assumptions rows 214-219; the share of members on longer terms is 0% until there is data, so projections use the $500 price. Unit economics and break-even use list price.
- Result at 5% churn: Base needs $421K, Year 3 profit $70K, investor not repaid within 5 years (0.47x); Conservative needs $617K. At 3% churn Base Year 3 profit is about $173K and payback is Year 5.
- Member cap stays 133, but at 5% churn the Upside case levels off near 120 and does not reach it within 36 months.

## Nov 11, 2026 (quarterly churn, winter sign-ups)
- Churn clarified as 5% per quarter (the 3-month term length), replacing 5% monthly. Cases: Conservative 6%, Base 5%, Upside 4% per quarter (the 1-point spread is an assumption), converted to a monthly equivalent in the engine (Assumptions rows 23-24).
- Sign-ups weighted toward winter because the facility is indoor (Assumptions B137:B148): Jan 1.4, Feb 1.3, Mar 1.1, Apr 0.9, May 0.8, Jun 0.7, Jul 0.7, Aug 0.8, Sep 0.9, Oct 1.0, Nov 1.1, Dec 1.3. Average 1.0. These factors are an assumption to confirm.
- Result: Base needs $403K, Year 3 profit $256K, first profitable month 5, investor repaid Year 4 (1.28x in 5 years). Conservative needs $446K (covered by the $500K), first profitable month 15, 0.64x in 5 years. Upside reaches the 133 cap and repays in Year 3. Buffer in the raise: $97K Base, $54K Conservative.

## Nov 11, 2026 (roles, key person)
- Slide 13 roles now include the GM (day-to-day operations, member experience, social media) and traditional investor roles (financial review, approval of major decisions, strategic sounding board). Counsel to confirm.
- Slide 14 key person risk now includes a term key person life insurance policy on Jeremy, owned by the business. Quote pending; premium not yet in the model.

## Nov 11, 2026 (Tonal quote)
- Tonal 2 quote received (Hannah Rosa, valid to 12/31/2026). Chosen: 5-year service term. Per unit: $4,685 equipment and delivery plus $1,445 a year service. Due at signing $6,130 per unit ($18,390 for 3); then $1,445 per unit a year for years 2 to 5 ($4,335 a year for 3, spread monthly in Assumptions B41). Five-year cost per unit $11,910, $35,730 for 3, against the earlier $33K estimate. Taxes not included.
- Gym floor is now $40,390 at opening (was $55K). Result: Base needs $390K, Conservative $437K, Base Year 3 profit $253K, first profitable month 6 (Conservative 16), investor repaid Year 4 (1.27x), break-even about 45 members.

## Nov 13, 2026 (hitting-station software)
- Software subscription is $1,200 a year per hitting station, 5 stations = $6,000 a year ($500 a month), added to monthly operating costs (Assumptions B41 = $700 base software and billing + $500 station software + Tonal service). The $700 base may already include some launch-monitor software; if so this double counts and is conservative. Confirm the breakdown.
- Result: Base needs $394K, Conservative $446K (the raise still covers both; buffer $106K / $54K). Base Year 3 profit $247K, first profitable month 6 (Conservative 16), investor repaid Year 4 (1.26x), Conservative 0.57x. Break-even about 46 members, fixed costs about $20,000 a month.

## Nov 13, 2026 (deck consistency pass)
- Confirmed: the $1,200 per station software is on top of the $700 base software and billing line.
- Whole-deck check against docs/NUMBERS.md: all numbers on slides 6, 9-13 match. Cover now says "Performance Club" (was "Hub"). Slide 5 no longer says "state-of-the-art" or "structured" testing.

## Nov 13, 2026 (opening date)
- No opening date is stated anywhere in the deck or kit; the date is not confirmed. Re-add only once it is firm.

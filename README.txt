VOTERS_LOCAL — full voter lists, families and GenZ, for this computer only.

These files hold every voter's name, age, relation, house number and EPIC.
Do not upload them, push them to GitHub, or put them on a website.

Open index.html in a browser (straight from disk; no server needed).
  index.html#186            Amethi, all-booth summary
  index.html#186/45/genz    Amethi booth 45, GenZ list

Rebuild (from /home/sayantan/CPIM/UP):
  ./up_family.py 186_Amethi_xlsx            # families  -> FAMILY/
  ./build_voters_local.py 186_Amethi        # this folder

# Invitee Parser Script for Tau Beta Pi
This script will take in a list of current members as well as the list of
eligible members and output a `csv` file of new members to invite to TBP

## Running
To run the script, make sure you have a file called `invitees.xlsx` which
contains the list of eligible members, and another file called `members.xlsx`
which contains the emails of every existing member.

Then run `$ python new_member_script.py` and your script will be complete!

## Format of Files
Your `xlsx` files must be formatted in this manner

### `invitees.xlsx`
This file should come from the Dean's office.  Check that it is formatted in
the following matter:

First Name | Last Name | Year | Email           |
-------------------------------------------------
Conner     | Chyung    | U4   | cchyung@usc.edu |
...        | ...       | ...  | ...             |


### `members.xlsx`
This file should just be a list of emails of current members:

Email           |
-----------------
cchyung@usc.edu |
...             |

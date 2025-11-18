from source.exel_styles import users_report
from source.block_functions import block_inactive
from source.config_loader import args


# The main function, which executes the script of gathering info from DB,
# then compares info to the table with filtered data and saves data to the
# Excel-file with the custom styles
if __name__ == "__main__":

    if args.user_report:
        users_report()

    if args.block_inactive:
        block_inactive()

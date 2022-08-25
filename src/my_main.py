import getmail
import str_to_xls
import xls_to_img
import sendmail_img

def main():
    str_to_xls.main()
    xls_to_img.main()
    sendmail_img.main()


main()

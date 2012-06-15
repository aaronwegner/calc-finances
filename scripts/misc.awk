#!/usr/bin/awk -f

function string_to_hex_string(str)
{
  begin = ""
  while (length(str) > 0) {
    c = substr(str, 1, 1)
    cmd = "printf \"0x%02x \" \"'"c"\""
    #print cmd
    cmd | getline more
    begin = begin more
    close(cmd)
    #str = substr(str, RSTART + RLENGTH)
    str = substr(str, 2)
  }
  return begin
}

function dollar_to_number(dollar)
{
  if (match(dollar, /[0-9,]*\.[0-9][0-9]/) > 0) {
    if ((dollar ~ /\(/) && (dollar ~ /\)/)) {
      # negative
      str = "-"substr(dollar, RSTART, RLENGTH)
    } else {
      # positive
      str = substr(dollar, RSTART, RLENGTH)
    }
  }
  return str
}

function file_name_get_line_terminators(file_name)
{
  # default to unix
  record_separator = "\n"
  verbose = 0
  #verbose = 1

  # try to obtain the line terminators of a file using the `file` command line utility
  cmd = "test -f "file_name" && echo is a file"
  cmd | getline tf
  close(cmd)
  if (tf == "") {
    if (verbose) {
      printf("%s is not a file\n", file_name)
    }
  } else {
    if (verbose) {
      printf("%s is a file\n", file_name)
    }
    # try to obtain the line terminators of a file using the `file` command line utility
    cmd = "file "file_name
    while ((cmd | getline) > 0) {
      if (verbose) {
        print $0
      }
      str = $0
      # seek for string " with ", then grap the next word of alphabetic characters
      if (match(str, / with /) > 0) {
        str = substr(str, RSTART + RLENGTH)
        if (match(str, /[[:alpha:]]+/) > 0) {
          str = substr(str, 1, RLENGTH)
          if (str == "CRLF") {
            # dos
            record_separator = "\r\n"
          } else if (str == "CR") {
            # mac
            record_separator = "\r"
          }
        }
      }
    } 
    # close this cmd before running string_to_hex_string as it also uses cmd and the
    # while loop will execute cmd a second time
    close(cmd)
  }

  if (verbose) {
    print "returning record separator " string_to_hex_string(record_separator)
  }
  return record_separator

  # here are some other methods tested but the above is better for now

  # cmd = "awk 'BEGIN { RS = \"\\n\" } {s = $0; i = 0; while (match(s, /\\r/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
  # cmd | getline
  # print "cr "$0
  # close(cmd)
  # cmd = "awk 'BEGIN { RS = \"\\r\" } {s = $0; i = 0; while (match(s, /\\n/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
  # cmd | getline
  # print "nl "$0
  # close(cmd)
  #
  # cmd = "awk 'BEGIN { RS = \"\\r\\n\" } {s = $0; i = 0; while (match(s, /\\r/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
  # cmd | getline
  # print "dos cr "$0
  # close(cmd)
  # cmd = "awk 'BEGIN { RS = \"\\r\\n\" } {s = $0; i = 0; while (match(s, /\\n/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
  # cmd | getline
  # print "dos nl "$0
  # close(cmd)

# this is the way to do it if AWK has more modern features, you can use match(s, r [, a]) instead of just match(s, r)

    # array[0] whole regular expression
    # array[1] first set of parentheses
    # array[2] second set of parentheses (the text we want)
    # array[3] third set of parentheses
#    if (match($0, /(.* with )([[:alpha:]]+)( .*)/, array) > 0) {
#      if (verbose) {
#        for (i = 0; i < length(array) / 3; i++) {
#          print "          array["i"] "array[i]
#          print " array["i", \"start\"] "array[i, "start"]
#          print "array["i", \"length\"] "array[i, "length"]
#        }
#      }
#      if (array[2] == "CRLF") {
#        # dos
#      } else if (array[2] == "CR") {
#        # mac
#      }
}

BEGIN {
  # if we're seeking a dollar amount dollar_token must be "" by definition
  dollar_token = ""
  dollar_seek = 0
  # do print is a workaround for Adobe Reader -> File -> Save as Text duplicates
  do_print = 1

  #print "    RS is "string_to_hex_string(RS)
	#for (i = 0; i < ARGC; i++) {
    #printf("ARGV[%d] is %s\n", i, ARGV[i])
  #}
  if (ARGC > 1) {
    # assume the last argument is a regular file
    RS = file_name_get_line_terminators(ARGV[ARGC - 1])
  #  print "RS is now "string_to_hex_string(RS)
  }
  if (print_header == 1) {
    printf("\"Transaction Date\",\"Description of Transaction\",\"Amount\"\n")
  }
}
{
  if (fixed < 1) {
    # default fixed is a date, for instance:
    # 2012/04/30 Followed by some sort of description of the transaction. $10.00
    # will be converted to:
    # "2012/04/30","Followed by some sort of description of the transaction.","10.00"
    fixed = 1
    #printf("fixed  is = %d\n", fixed)
  }
  if ((dollar_seek == 0) && ($1 ~ /[0-9][0-9]\/[0-9][0-9]/)) {
    if (dollar_token != "") {
      if (do_print) {
        print dollar_token
      }
      dollar_token = ""
    }
    dollar_index = -1
    # helper means the token before the dollar token, or the last token in a line
    # it is always interpreted to be a part of the column with a variable number of tokens
    # or a part of the description of the transaction in other words
    helper_index = 1
    for (i = 2; i <= NF; i++) {
      if ($i ~ /[(]*\$[0-9,]*\.[0-9][0-9][)]*/) {
        # we found the dollar token, and set the fixed index to the one before
        dollar_index = i
        helper_index = i - 1
        # this token is a dollar amount
        dollar_seek = 0
        #dollar_token="\" \""$NF"\""
        dollar_token="\",\""dollar_to_number($i)"\""
        break
      }
    }
    if (i > NF) {
      helper_index = NF
    }
    if (dollar_index == -1) { 
      # This is a rare case in which the first line does not contain a dollar amount.  Usually,
      # even on a multi-line transaction, we see the dollar amount at the end of the line:
      #
      # 10/16 10/16 852185392011SHWX6 PAYPAL PURCH, 4029357733 SAN JOSE CA $188.31 
      # ETSY, INC 
      #
      # However, in the 20111025.txt file we see one instance of the dollar amount not
      # being at the end of the line:
      #
      # 10/07 10/07 55457028T600K4FEP MEDITERRANEAN KEBOB HO DELRAY BEACH 
      # FL 
      # $26.35 
      #
      # So, we need to account for this, apparently.
      #
      dollar_seek = 1
      dollar_token = ""
    }
    #if (!a[$0]++) {
    #  do_print = 1
    #} else {
    #  do_print = 0
    #}
    if (do_print) {
      #printf("\"%s\",\"%s\",\"%s\",\"", $1, $2, $3)
      #printf("\"%s\",\"", $1)
      printf("\"%s\",", $1)
      for (i = 2; i <= fixed; i++) {
        if (i < helper_index) {
          printf("\"%s\",", $i)
        } else {
          printf("\"\",")
        }
      }

      # fixed = 3
      #
      # 10/08 10/08 05436848S002GNXXV TARGET $21.61
      # 10/08 10/08 05436848S002GNXXV $21.61
      # 10/08 05436848S002GNXXV $21.61
      # 10/10 $28.86

      printf("\"")
      if ((helper_index != 1) && (helper_index <= fixed)) {
        # token at helper index (if not 1) always assumed to be part of variable number of tokens column
        printf("%s", $helper_index)
      }
      for (i = fixed + 1; i <= helper_index; i++) {
        if (i == helper_index) {
          # last token on line
          printf("%s", $i)
        } else {
          # not last token on line
          printf("%s ", $i)
        }
      }
    }
  } else if (dollar_seek == 1) {


    dollar_index = -1
    # here helper index is the last token we print on a line
    helper_index = NF
    for (i = 1; i <= NF; i++) {
      if ($i ~ /[(]*\$[0-9,]*\.[0-9][0-9][)]*/) {
        # we found the dollar token, and set the fixed index to the one before
        dollar_index = i
        helper_index = i - 1
        # this token is a dollar amount
        dollar_seek = 0
        #dollar_token="\" \""$NF"\""
        dollar_token="\",\""dollar_to_number($i)"\""
        break
      }
    }
    if (do_print) {
      if (dollar_index != 1) {
        # first token is not dollar amount, do a newline
        printf("\n")
      }
      for (i = 1; i <= helper_index; i++) {
        if (i == helper_index) {
          printf("%s", $i)
        } else {
          printf("%s ", $i)
        }
      }
    }
  } else {
    # either we're not seeking a dollar amount or the last token is not a dollar amount
    # print the whole line
    if (do_print) {
      printf("\n")
      for (i = 1; i <= NF; i++) {
        if (i == NF) {
          printf("%s", $i)
        } else {
          printf("%s ", $i)
        }
      }
    }
  }
}

END {
  if (do_print) {
    if (dollar_token != "") {
      print dollar_token
    }
  }
}

#  for (i=1; i<=NF; i++) {
#    if ($i ~ /[0-9]+\.[0-9][0-9]$/) {
#      print $i
#    }
#  }

#  cmd = "strip "$1
#  while ( ( cmd | getline result ) > 0 ) {
#    print  result
#  } 
#  close(cmd)

  #print "ARGC "ARGC
  #print "ARGV[0] "ARGV[0]
  #print "ARGV[1] "ARGV[1]
  #for (i = 1; i < ARGC; i++) {
    #print "ARGV["i"] " ARGV[i]
  #}


#    if (ARGC == 2) {
#      cmd = "awk 'BEGIN { RS = \"\\n\" } {s = $0; i = 0; while (match(s, /\\r/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
#      cmd | getline
#      print "cr "$0
#      close(cmd)
#      cmd = "awk 'BEGIN { RS = \"\\r\" } {s = $0; i = 0; while (match(s, /\\n/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
#      cmd | getline
#      print "nl "$0
#      close(cmd)
#
#      cmd = "awk 'BEGIN { RS = \"\\r\\n\" } {s = $0; i = 0; while (match(s, /\\r/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
#      cmd | getline
#      print "dos cr "$0
#      close(cmd)
#      cmd = "awk 'BEGIN { RS = \"\\r\\n\" } {s = $0; i = 0; while (match(s, /\\n/) > 0) { i++; s = substr(s, RSTART+RLENGTH) } } END { print i }' "ARGV[1]
#      cmd | getline
#      print "dos nl "$0
#      close(cmd)
#    }


#    cmd = "test -f "ARGV[i]"; echo $?"
#    cmd = "awk 'BEGIN{RS=\"\\n\"}{numrecords=FNR}END{print numrecords}' "ARGV[1]
#    cmd | getline
#    unix = $0
#    close(cmd)
#    cmd = "awk 'BEGIN{RS=\"\\r\"}{numrecords=FNR}END{print numrecords}' "ARGV[1]
#    cmd | getline
#    mac = $0
#    close(cmd)
#    cmd = "awk 'BEGIN{RS=\"\\r\\n\"}{numrecords=FNR}END{print numrecords}' "ARGV[1]
#    cmd | getline
#    dos = $0
#    close(cmd)
#    print "unix is "unix" mac is "mac" dos is "dos
#    if ((unix > dos) && (unix > mac)) {
#      print "guess unix"
#    } else if ((mac > unix) && (mac > dos)) {
#      print "guess mac"
#    } else if (dos == unix) {
#      print "guess dos"
#    } else {
#      print "default guess unix"
#    }

#BEGIN {
#  printf("timeout 9999\n");
#}
#
#{
#  if ($1 == "cp.b" || $1 == "erase") {
#    printf("send \"%s\"\nexpect {\n  \"\\n=> \"\n  timeout 9999\n}\n", $0);
#  } else if ($1 == "reset") {
#    printf("send \"%s\"\nexpect {\n  \"autoboot:\"\n}\n", $0);
#    printf("send \"A\"\nexpect {\n  \"\\n=> \"\n}\n");
#  } else {
#    printf("send \"%s\"\nexpect {\n  \"\\n=> \"\n}\n", $0);
#  }
#}

#BEGIN {
#  object = ARGV[1];
#  ARGV[1] = "";

#  while ("nm " object "| sort" | getline) {
#    if ($2 == "t" || $2 == "T") {
#      address[i] = "0x" $1; name[i] = $3;
#      i++;
#    }
#  }
#  syms = i;
#}

#{
#  lo = 0;
#  hi = syms - 1;

#  while ((hi-1) > lo)
#    {
#      try = int ((hi + lo) / 2);
#      if ($0 < address[try])
#	hi = try;
#      else if ($0 >= address[try])
#	lo = try;
#    }
#  print name[lo] "\n"; fflush();
#}

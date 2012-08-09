# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

BEGIN { $| = 1; print "1..4\n"; }
END {print "not ok 1\n" unless $loaded;}

#use strict;
use Data::Table;
use Data::Table::Excel qw(tables2xls xls2tables tables2xlsx xlsx2tables);
use Data::Dumper;

$loaded = 1;
print "ok loaded\n";

# read in two CSV tables and generate an Excel file with two spreadsheets
my $t_category = Data::Table::fromFile("Category.csv");
#my $s_A = $t_A->csv;
my $t_product = Data::Table::fromFile("Product.csv");
#my $s_B = $t_B->csv;
if (-e 'NorthWind.xls') { unlink("NorthWind.xls") }
tables2xls("NorthWind.xls", [$t_category, $t_product], ["Category","Product"]);

if (-e 'NorthWind.xls') {
  print "ok 1 table2xls\n";
} else {
  print "not ok 1 table2xls\n";
}

# read in TableAB.xls file as two Data::Table objects
my ($tableObjects, $tableNames)=xls2tables("NorthWind.xls");
if (scalar @$tableObjects !=2 || scalar @$tableNames !=2) {
  print "not ok 2 xls2tables\n";
} else {
#print Dumper($tableObjects);
#print Dumper($tableNames);
  if ($tableNames->[0] eq 'Category' && $tableNames->[1] eq 'Product' &&
    $tableObjects->[0]->nofRow==$t_category->nofRow && $tableObjects->[1]->nofRow==$t_product->nofRow &&
    $tableObjects->[0]->nofCol==$t_category->nofCol && $tableObjects->[1]->nofCol==$t_product->nofCol) {
    print "ok 2 xls2tables\n";
  } else {
    print "not ok 2 xls2tables\n";
  }
}

if (-e 'NorthWind.xlsx') { unlink("NorthWind.xlsx") }
tables2xlsx("NorthWind.xlsx", [$t_category, $t_product], undef, [['silver','white','black'], [45,'white',37]]); # ["Category","Product"]);

if (-e 'NorthWind.xlsx') {
  print "ok 3 table2xlsx\n";
} else {
  print "not ok 3 table2xlsx\n";
}

# read in TableAB.xls file as two Data::Table objects
my ($tableObjects, $tableNames)=xlsx2tables("NorthWind.xlsx");
if (scalar @$tableObjects !=2 || scalar @$tableNames !=2) {
  print "not ok 4 xlsx2tables\n";
} else {
#print Dumper($tableObjects);
#print Dumper($tableNames);
  if ($tableNames->[0] eq 'Sheet1' && $tableNames->[1] eq 'Sheet2' &&
    $tableObjects->[0]->nofRow==$t_category->nofRow && $tableObjects->[1]->nofRow==$t_product->nofRow &&
    $tableObjects->[0]->nofCol==$t_category->nofCol && $tableObjects->[1]->nofCol==$t_product->nofCol) {
    print "ok 4 xlsx2tables\n";
  } else {
    print "not ok 4 xlsx2tables\n";
  }
}

1;

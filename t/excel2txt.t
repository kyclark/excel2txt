use strict;
use Config;
use Test::More 'no_plan';
use FindBin qw( $Bin );
use File::Path;
use File::Spec::Functions;
use File::Temp qw( tempdir );

require_ok( 'Spreadsheet::ParseExcel' );

my $perl        = $Config{'perlpath'};
my $excel2txt   = "$Bin/../excel2txt";
my $spreadsheet = "$Bin/test.xls";

ok( -e $excel2txt, "Script '$excel2txt' exists");
ok( -e $spreadsheet, "Spreadsheet '$spreadsheet' exists");

my $tempdir = tempdir( CLEANUP => 1 );
ok( chdir $tempdir, "Changed dir to '$tempdir'" );

my $res = `$perl $excel2txt $spreadsheet`;

ok( $res =~ /^Done/, "Processed OK" );

my $people   = catfile( $tempdir, 'People.txt' );
my $salaries = catfile( $tempdir, 'Salaries.txt' );
my $empty    = catfile( $tempdir, 'Empty.txt' );
ok( -e $people, "People spreadsheet converted to '$people'" );
ok( -e $salaries, "Salaries spreadsheet converted to '$salaries'" );
ok( !-e $empty, "Empty spreadsheet NOT converted to '$empty'" );

open my $pfh, '<', $people or die $!;
my @people_data = <$pfh>;
chomp @people_data;

cmp_ok( scalar @people_data, '==', 4, "Four lines in '$people'" );
cmp_ok( $people_data[0], 'eq', join("\t", qw[ first_name last_name title ]), 
    'Headers OK');

open my $sfh, '<', $salaries or die $!;
my @salary_data = <$sfh>;
chomp @salary_data;

cmp_ok( scalar @salary_data, '==', 5, "Five lines in '$salaries'" );

ok( rmtree( $tempdir ), "Removed $tempdir" );

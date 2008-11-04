use strict;
use Config;
use Test::More tests => 31;
use FindBin qw( $Bin );
use File::Basename qw( basename );
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

for my $format ( qw[ txt html yaml xml ] ) {
    my $people   = catfile( $tempdir, "test-people.$format" );
    my $salaries = catfile( $tempdir, "test-salaries.$format" );
    my $empty    = catfile( $tempdir, "test-empty.$format" );

    my @res = split /\n/, 
        `$perl $excel2txt -f $format -n -o $tempdir $spreadsheet`;

    ok( $res[-1] =~ /^Done/, "Processed format '$format' OK" );
    ok( -e $people, 
        sprintf("People spreadsheet converted to '%s'", basename($people))
    );
    ok( -e $salaries, 
        sprintf("Salaries spreadsheet converted to '%s'", basename($salaries))
    );
    ok( !-e $empty, 
        sprintf("Empty spreadsheet NOT converted to '%s'", basename($empty))
    );
}

my $people = catfile( $tempdir, 'test-people.txt' );
open my $pfh, '<', $people or die $!;
my @people_data = <$pfh>;
chomp @people_data;

cmp_ok( scalar @people_data, '==', 4, 'Four lines in people text file' );
cmp_ok( $people_data[0], 'eq', join("\t", qw[ first_name last_name title ]), 
    'Headers OK');

my $salaries = catfile( $tempdir, 'test-salaries.txt' );
open my $sfh, '<', $salaries or die $!;
my @salary_data = <$sfh>;
chomp @salary_data;

cmp_ok( scalar @salary_data, '==', 5, 'Five lines in salaries text file' );

my $html_file = catfile( $tempdir, 'test-people.html' );
open my $html, '<', $html_file or die $!;
my @lines = <$html>;
is( scalar @lines, 22, 'HTML file is correct length' );

SKIP: {
    eval { require XML::Simple };
    if ( $@ ) {
        skip "Can't find XML::Simple" => 1;
    }

    my $xml_file = catfile( $tempdir, 'test-people.xml' );
    my $xml      = XML::Simple::XMLin( $xml_file );
    is( ref $xml, 'HASH', 'XML loaded as hash' );
    is( scalar @{ $xml->{'people'} }, 3, 'Found three people' );
    is( ref $xml->{'people'}[2], 'HASH', 'Records is a hash' );
    is( $xml->{'people'}[2]->{'first_name'}, 'Geoffrey', 'Found Geoff' );
}

SKIP: {
    eval { require YAML };
    if ( $@ ) {
        skip "Can't find YAML" => 1;
    }

    my $yaml_file = catfile( $tempdir, 'test-people.yaml' );
    my @data      = YAML::LoadFile( $yaml_file ); 
    is( scalar @data, 3, 'YAML loaded 3 records' );
    is( ref $data[1], 'HASH', 'Record 2 is a HASH' );
    is( $data[1]->{'title'}, 'Vice-President', 'Record 2 title is VP' );
}

ok( rmtree( $tempdir ), 'Removed temp dir' );

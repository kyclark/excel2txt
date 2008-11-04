use Module::Build;

my $build = Module::Build->new(
    dist_name      => 'excel2txt',
    dist_version   => '0.04',
    dist_author    => 'Ken Youens-Clark <kclark@cpan.org>',
    dist_abstract  => 'convert Excel data to text formats',
    license        => 'perl',
    requires       => {
        'Pod::Usage'              => 0,
        'Readonly'                => 0,
        'Spreadsheet::ParseExcel' => 0.2603,
    },
    recommends     => {
        'YAML'                    => 0,
        'XML::Simple'             => 0,
    },
    build_requires => {
        'Test::More'              => 0,
    },
    script_files   => [ 'excel2txt' ],
);
$build->create_build_script;

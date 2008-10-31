use Module::Build;

my $build = Module::Build->new(
    dist_name      => 'excel2txt',
    dist_version   => '0.01',
    dist_author    => 'Ken Youens-Clark <kclark@cpan.org>',
    dist_abstract  => 'convert Excel data to delimited text files',
    license        => 'perl',
    requires       => {
        'Pod::Usage'              => 0,
        'Readonly'                => 0,
        'Spreadsheet::ParseExcel' => 0.2603,
    },
    build_requires => {
        'Test::More'              => 0,
    },
    script_files   => [ 'excel2txt' ],
);
$build->create_build_script;

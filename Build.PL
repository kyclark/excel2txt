use Module::Build;

my $build = Module::Build->new(
    dist_name         => 'excel2txt',
    dist_version_from => 'excel2txt',
    dist_author       => 'Ken Youens-Clark <kclark@cpan.org>',
    dist_abstract     => 'convert Excel data to text formats',
    license           => 'perl',
    script_files      => [ 'excel2txt' ],
    requires          => {
        'Pod::Usage'              => 0,
        'Readonly'                => 0,
        'Spreadsheet::ParseExcel' => 0.2603,
    },
    recommends        => {
        'XML::Simple'             => 0,
        'YAML'                    => 0,
    },
    build_requires    => {
        'Test::More'              => 0,
    },
);

$build->create_build_script;

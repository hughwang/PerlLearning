use strict;
use warnings;
use diagnostics;
use Template::Constants qw( :debug );
use Template;
use StringUtils;

use HTTP::Request::Common;
use LWP;
use Scalar::Util qw(reftype);
use HTTP::Cookies::Netscape;
use DBI;
use Sql;


use utf8;
use IO::Handle;
use Carp;
use Net::Google::Spreadsheets;

my $outdir = "NJ_tmp";
my $cookie_file = "NJ_tmp/cookiejar.txt";
system("mkdir -p $outdir") if (! -d $outdir);

my $agent = &do_login();

our $db_string = "K:\\eproj\\trunk\\USPS_Route\\DM.accdb";

our $dbh = DBI->connect("DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=$db_string");
unless($dbh) {
	print "error open database: $DBI::errstr\n";
	exit;
}
my $sql= "SELECT ZIP,CRRT from DirectMailRoute";
	 
my $sth1 = $dbh->prepare($sql)
				or die "Couldn't prepare statement: " . $dbh->errstr;

$sth1->execute()             # Execute the query
			or die "Couldn't execute statement: " . $sth1->errstr;

while (my @data = $sth1->fetchrow_array()) {
	my $Zip= $data[0];
	my $Route = $data[1];

	
   
    
    my $mapurl;
	# Output file
    my $of;

    $mapurl = "http://www.melissadata.com/lookups/MapCartS.asp?zip=$Zip&cart=$Route&mp=f";
    $of = "$outdir/${Zip}_${Route}.html";
	
    if( ! -e $of) {
        print STDERR "$mapurl\n";  
        &get_page($of, $agent, $mapurl);
        sleep(20);
    }
    my $housevalue = ConvertFile($of, $Zip,$Route);
    my $file_name = "${Zip}_${Route}.htm";
}

print STDERR "  Done\n";
$dbh->disconnect;

exit(0);

sub ConvertFile {
    my ($inputFile, $Zip,$Route) = @_;
    my $tt= Template->new({
        #DEBUG => DEBUG_PARSER | DEBUG_PROVIDER | DEBUG_CONTEXT|DEBUG_DIRS,
            INCLUDE_PATH => "templates",
    });    
    my $vars;
    
    open IN, "<$inputFile" || croak "$inputFile can not be opened";
    
    my @lines=<IN>;
    close(IN);
    my @polygons = ();
    my $polygon_ref;
    my $point_ref;
    my $center_lat;
    my $center_long;
    
    my $address_lat=undef;
    my $address_long;
    my $address;
    
    my $Riterator = &StringUtils::GetListIterator( \@lines);
    
    my $line;
    my $housevalue=0;
    
    
    my $count = 0;
    my $total_lat=0;
    my $total_long=0;
    PROCESS_NEW_LINE:
    while ( &$Riterator( \$line ) ) {
        if($line=~ /var points =/) {
            $polygon_ref=();
        } elsif( $line=~ /points.push\(new Microsoft.Maps.Location\((.+?),(.+?)\)\)/) {
            $count=$count+1;
            $total_lat = $total_lat+ $1;
            $total_long = $total_long+ $2;
    
            $point_ref = {
                       Latitude => $1,
                       Longitude=>$2
        };
            push @{$polygon_ref},$point_ref;
        } elsif($line=~ /var polygon =/) {
            push @polygons,$polygon_ref;
        #} 
        #elsif($line=~ /Microsoft.Maps.LocationRect.fromString\('(.+?),(.+?),(.+?),(.+?)'\)/) {
        #    my $left_up_lat = $1;
        #    my $left_up_long = $2;
        #    my $right_bottom_lat = $3;
        #    my $right_bottom_long = $4;
        #    $center_lat = ($left_up_lat + $right_bottom_lat) /2;
        #    $center_long = ($left_up_long + $right_bottom_long) /2;  
        } elsif($line=~ /Maps.Pushpin\(new Microsoft.Maps.Location\((.+?)\s*,(.+?)\s*\)/) {
            next PROCESS_NEW_LINE if($address_lat);
            $address_lat = $1;
            $address_long = $2;
            &$Riterator( \$line );
            &$Riterator( \$line );
            if($line=~ /pin.Description = '(.+?)'/) {
                $address = $1;
                $address =~ s/to \d+//;
                $address =~ s/ \(.*\)//;
                $address =~ s/<br>/,/;
                $address =~ s/<br>/ /;
            }
        } elsif($line=~ /Average Home Value<\/strong><\/td><td align=right style='padding-right:5px'>\$(.+?)</) {
            $housevalue = $1;
        }
        #var pin = new Microsoft.Maps.Pushpin(new Microsoft.Maps.Location(38.930699 ,-77.278358 ),{text: '61'} );
        #pin.Title = '<b>ZIP+4 Code 22182-1922</b>';
        #pin.Description = '1701 to 1799 (Odd)    BROOKSIDE LN   <br>VIENNA, VA<br>22182-1922';
    
        
        
        
    }
	if( !$count) {
		return 0;
	}
    $center_lat = $total_lat/$count;
    $center_long = $total_long/$count;
    
    my $output = "output/${Zip}_${Route}.htm";
    if(!$Route) {
        $output = "output/${Zip}.htm";
    }
    my $template = 'drawPolygon.tt';
    
    my $new_persecution_type_name= 'test';
    $vars = {
        'Polygons' => \@polygons,
        'Center_Latitude' => $center_lat ,
        'Center_Longitude' => $center_long ,
        'Address_Latitude' => $address_lat ,
        'Address_Longitude' => $address_long ,
        'Address' => $address,
        Zip => $Zip,
        Route => $Route,
    };
    #$tt->process($template, $vars,$output);
    #print Dumper($vars);
    $tt->process($template, $vars,$output);
    return $housevalue;
}

# ==================================

sub get_page
{
        my ($ofile, $agent, $url) = @_;
        &save_page($ofile, $agent, "GET", $url);
}

sub post_page
{
        my ($ofile, $agent, $url, $data) = @_;
        &save_page($ofile, $agent, "POST", $url, $data);
}

sub save_page
{
        my ($of, $agent, $method, $url, $data) = @_;

        my $request;
        if ($method =~ /^POST$/io) {
                $request  = HTTP::Request::Common::POST($url, $data);
        } elsif ($method =~ /^GET$/io) {
                $request  = HTTP::Request::Common::GET($url);
        } else {
                return;
        }

        $request->header('User_Agent' => 'Mozilla/5.0 (Windows NT 5.1; rv:11.0) Gecko/20100101 Firefox/11.0');
        $request->header('Accept' => 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8');

        my $response = $agent->request($request);


        my %FD;

        open($FD{$of}, ">$of");
        my $fd = $FD{$of};

#       print $fd "_request=";
#       &pr_response($fd, 1, $response->{_request});
#       print $fd "\n";

#       #print $fd "_uri=", $response->{_uri}, "\n";
#       #print $fd "method=", $response->{_method}, "\n";

#       print $fd "_headers=";
#       &pr_response($fd, 0, $response->{_headers});
#       print $fd "\n";
#       print $fd "_content=", $response->{_content};

        print $fd $response->{_content};

        close($fd);
}

sub pr_response
{
        my ($fd, $level, $response) = @_;

        foreach my $key (sort keys %{$response}) {
#               print $fd "  "x$level, "ref($key)=", reftype($response->{$key}), "\n";
                if (reftype($response->{$key}) eq "ARRAY") {
                        print $fd " "x(4*$level), "$key=", join("|", @{$response->{$key}}), "\n";
                } else {
                        print $fd " "x(4*$level), "$key=$response->{$key}\n";
                }
                if (reftype($response->{$key}) eq "HASH") {
                        pr_response($fd,$level+1, $response->{$key});
                }
        }
}


sub do_login
{
        my $agent = LWP::UserAgent->new;

        #unlink $cookie_file;
        my $cookie_jar = HTTP::Cookies::Netscape->new(file => $cookie_file, autosave => 1, ignore_discard => 1, hide_cookie2 => 1);
        $agent->cookie_jar($cookie_jar);

        my $user = "jimmy2007\@edoors.com";
        my $passwd = "jimmypwd";
        my $loginurl = 'https://www.melissadata.com/user/signin.aspx?src=http%3a%2f%2fwww.melissadata.com%2flookups%2fmapzipv.asp';


        my $key_user = 'ctl00$ContentPlaceHolder1$Signin1$UserLogin$UserName';
        my $key_passwd = 'ctl00$ContentPlaceHolder1$Signin1$UserLogin$Password';
        my $key_login_button = 'ctl00$ContentPlaceHolder1$Signin1$UserLogin$LoginButton';

        my %data = (
        $key_user, $user,
        $key_passwd, $passwd,
        $key_login_button, "Sign%20In",
        '__EVENTTARGET', "",
        );

        my $of = "$outdir/signin.html";
        # use same login session
        if (-M $cookie_file < 1 && -s $cookie_file > 50 && -s $of > 1000) {
                $cookie_jar->load($cookie_file);
        } else {
                print STDERR "Signin\n";
                &post_page($of, $agent, $loginurl, \%data);
                print STDERR "Cookie=", $cookie_jar->as_string(!$cookie_jar->{ignore_discard}), "\n";
                $cookie_jar->save();
        }
        return $agent;
}





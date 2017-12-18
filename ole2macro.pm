# <@LICENSE>
# Licensed to the Apache Software Foundation (ASF) under one or more
# contributor license agreements.  See the NOTICE file distributed with
# this work for additional information regarding copyright ownership.
# The ASF licenses this file to you under the Apache License, Version 2.0
# (the "License"); you may not use this file except in compliance with
# the License.  You may obtain a copy of the License at:
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# </@LICENSE>

=head1 NAME

OLE2Macro - Look for Macro Embedded Microsoft Word and Excel Documents

=head1 SYNOPSIS

 loadplugin     ole2macro.pm
 body MICROSOFT_OLE2MACRO eval:check_microsoft_ole2macro()
 score MICROSOFT_OLE2MACRO 4

=head1 DESCRIPTION

Detects embedded OLE2 Macros embedded in Word and Excel Documents. Based on:
https://blog.rootshell.be/2015/01/08/searching-for-microsoft-office-files-containing-macro/

 10/12/2015 - Jonathan Thorpe - jthorpe@conexim.com.au
 08/11/2017 - Giovanni Bechis - Complete rewrite based on OLE::Storage_Lite

=head1 USER SETTINGS

=over 4

=item archived_files_limit n	(default: 3)

How many files within an archive we should process

=back

=over 4

=item file_max_read_size n	(default: 102400)

How much amount of bytes we can read from a file

=back

=cut

package OLE2Macro;

use Mail::SpamAssassin::Plugin;
use Mail::SpamAssassin::Logger;
use Mail::SpamAssassin::Util;
use File::Temp qw/ :POSIX /;
use IO::Uncompress::Unzip;
use MIME::Parser;
use OLE::Storage_Lite;

use strict;
use warnings;
use bytes;
use re 'taint';
use vars qw(@ISA);

@ISA = qw(Mail::SpamAssassin::Plugin);

#File types and markers
my $match_types = qr/(?:word|excel)$/;
my $match_types_ext = qr/(?:doc|dot|xls)$/;

#Microsoft OOXML-based formats with Macros
my $match_types_xml = qr/(?:xlsm|xltm|xlsb|potm|pptm|ppsm|docm|docx|dotm)$/;

sub new;
sub set_config;
sub check_microsoft_ole2macro;
sub _check_mail;
sub _check_attachment;
sub _check_OLE;

# constructor: register the eval rule
sub new {
    my $class = shift;
    my $mailsaobject = shift;

   # some boilerplate...
    $class = ref($class) || $class;
    my $self = $class->SUPER::new($mailsaobject);
    bless ($self, $class);

    $self->register_eval_rule("check_microsoft_ole2macro");

    return $self;
}

sub set_config {

  my ($self, $conf) = @_;
  my @cmds;

  push(@cmds, {
    setting => 'archived_files_limit',
    default => 3,
    type => $Mail::SpamAssassin::Conf::CONF_TYPE_NUMERIC,
  });

  push(@cmds, {
    setting => 'file_max_read_size',
    default => 102400,
    type => $Mail::SpamAssassin::Conf::CONF_TYPE_NUMERIC,
  });

  $conf->{parser}->register_commands(\@cmds);
}


sub check_microsoft_ole2macro {
    my ($self, $pms) = @_;

    _check_mail(@_) unless exists $pms->{nomacro_microsoft_ole2macro};
    return $pms->{nomacro_microsoft_ole2macro};
}

sub _check_mail {
   my ($self, $pms) = @_;

   my $processed_files_counter = 0;
   $pms->{nomacro_microsoft_ole2macro} = 0;

   my $fullref = \$pms->get_message()->get_pristine();
   my $parser = MIME::Parser->new();
   $parser->output_to_core(1);      # don't write attachments to disk
   my $message  = $parser->parse_data($fullref);

   foreach my $part ($message->parts_DFS) {
    my $content_type = $part->effective_type;
    my $body         = $part->bodyhandle;
    if ($content_type =~ $match_types) {
        _check_attachment($pms, $body);
    }
    if (($content_type =~ /application\/zip/) || ($content_type =~ /application\/vnd.openxml/)) {
        my $contents = $body->as_string;
        my $z = new IO::Uncompress::Unzip \$contents;

        my $status;
        my $buff;
        my $zip_fn;
	my $archived_files_process_limit = $self->{main}->{conf}->{archived_files_limit};
	my $file_max_read_size = $self->{main}->{conf}->{file_max_read_size};

        if (defined $z) {
            for ($status = 1; $status > 0; $status = $z->nextStream()) {
                $zip_fn = lc $z->getHeaderInfo()->{Name};

                #Parse these first as they don't need handling of the contents.
                if ($zip_fn =~ $match_types_xml) {
                    $pms->{nomacro_microsoft_ole2macro} = 1;
                    last;
                } elsif ($zip_fn =~ $match_types_ext or $zip_fn eq "[content_types].xml") {
                    $processed_files_counter += 1;
                    if ($processed_files_counter > $archived_files_process_limit) {
                        dbg( "Stopping processing archive on file ".$z->getHeaderInfo()->{Name}.": processed files count limit reached\n" );
                        last;
                    }
                    my $attachment_data = "";
                    my $read_size = 0;
                    while (($status = $z->read( $buff )) > 0) {
                        $attachment_data .= $buff;
                        $read_size += length( $buff );
                        if ($read_size > $file_max_read_size) {
                            dbg( "Stopping processing file ".$z->getHeaderInfo()->{Name}." in archive: processed file size overlimit\n" );
                            last;
                        }
                    }

                    #OOXML format
                    if ($zip_fn eq "[content_types].xml") {
                        if($attachment_data =~ /ContentType=["']application\/vnd.ms-office.vbaProject["']/i){
                            $pms->{nomacro_microsoft_ole2macro} = 1;
                            last;
                        }
                    } else {
                            _check_attachment($pms, $attachment_data);
                    }
                }
            }
        }
    }
   }
}

sub _check_attachment {
	my($pms, $body) = @_;
	my $tmpname = tmpnam();
	open OUT, ">$tmpname";
	binmode OUT;
	if ( $body->can("as_string") ) {
		print OUT $body->as_string;
	} else {
		print OUT $body;
	}
	close OUT;
	my $oOl = OLE::Storage_Lite->new($tmpname);
	my $oPps = $oOl->getPpsTree();
	my $iTtl = 0;
	my $result = _check_OLE($pms, $oPps, 0, \$iTtl, 1);
	# dbg("OLE2: " . $pms->{nomacro_microsoft_ole2macro});
	if($pms->{nomacro_microsoft_ole2macro} eq 1) {
		last;
		return 1;
	}
	unlink($tmpname);
}

sub _check_OLE {
  my($pms, $oPps, $iLvl, $iTtl, $iDir) = @_;
  my %sPpsName = (1 => 'DIR', 2 => 'FILE', 5=>'ROOT');

  # Make Name (including PPS-no and level)
  my $sName = OLE::Storage_Lite::Ucs2Asc($oPps->{Name});
  $sName = sprintf("%s", $sName);
  if($sName eq "_VBA_PROJECT") {
	# dbg("OLE2: " . $sName);
	$pms->{nomacro_microsoft_ole2macro} = 1;
        return 1;
  }
# For its Children
  my $iDirN=1;
  foreach my $iItem (@{$oPps->{Child}}) {
    _check_OLE($pms, $iItem, $iLvl+1, $iTtl, $iDirN);
    $iDirN++;
  }
  return 0;
}

1;

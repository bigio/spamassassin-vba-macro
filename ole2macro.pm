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
08/11/2017 - Giovanni Bechis - Complete rewrite based in OLE::Storage_Lite

=back

=cut

package OLE2Macro;

use Mail::SpamAssassin::Plugin;
use Mail::SpamAssassin::Logger;
use Mail::SpamAssassin::Util;
use MIME::Parser;
use OLE::Storage_Lite;
use File::Temp qw/ :POSIX /;

use strict;
use warnings;
use bytes;
use re 'taint';

use vars qw(@ISA);
@ISA = qw(Mail::SpamAssassin::Plugin);

#File types and markers
my $match_types = qr/(?:word|excel)$/;

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

sub check_microsoft_ole2macro {
    my ($self, $pms) = @_;

    _check_attachments(@_) unless exists $pms->{nomacro_microsoft_ole2macro};
    return $pms->{nomacro_microsoft_ole2macro};
}

sub _check_attachments {
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
	my $tmpname = tmpnam();
	open OUT, ">$tmpname";
	binmode OUT;
	print OUT $body->as_string;
	close OUT;
	my $oOl = OLE::Storage_Lite->new($tmpname);
	my $oPps = $oOl->getPpsTree();
	my $iTtl = 0;
	my $result = check_OLE($pms, $oPps, 0, \$iTtl, 1);
	# dbg("OLE2: " . $pms->{nomacro_microsoft_ole2macro});
	if($pms->{nomacro_microsoft_ole2macro} eq 1) {
		last;
        }
	unlink($tmpname);
     }
   }
}

sub check_OLE($$\$$) {
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
    check_OLE($pms, $iItem, $iLvl+1, $iTtl, $iDirN);
    $iDirN++;
  }
  return 0;
}

1;

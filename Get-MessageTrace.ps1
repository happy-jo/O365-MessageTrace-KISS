{\rtf1\ansi\ansicpg1252\cocoartf1671
{\fonttbl\f0\fswiss\fcharset0 Helvetica;\f1\fswiss\fcharset0 ArialMT;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue0;\red255\green255\blue255;\red0\green0\blue233;
}
{\*\expandedcolortbl;;\cssrgb\c0\c0\c0\c84706;\cssrgb\c100000\c100000\c100000;\cssrgb\c0\c0\c93333;
}
\margl1440\margr1440\vieww32360\viewh14760\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 # Get-MessageTrace.ps1\
# The user conducting the MessageTrace query must be assigned the \'93Message Tracking \'93 role within Exchange online.\
# This is a api query and is not limited to powershelgl\
\
\pard\pardeftab720\sl360\partightenfactor0

\fs29\fsmilli14667 \AppleTypeServices\AppleTypeServicesF65539 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 $creds = Get-Credential\AppleTypeServices \'a0\

\f1\fs24 \cb1 \

\f0\fs29\fsmilli14667 \AppleTypeServices\AppleTypeServicesF65539 \cb3 $MTHeaderParams\'a0= @\{DataServiceVersion="2.0";MaxDataServiceVersion="2.0";'Accept-Language'="EN-US";Accept="application/json";'Content-Type'="application/json"\}\AppleTypeServices \'a0\

\f1\fs24 \cb1 \

\f0\fs29\fsmilli14667 \AppleTypeServices\AppleTypeServicesF65539 \cb3 $messageTrace\'a0= (Invoke-WebRequest\'a0-Headers $MTHeaderParams\'a0-Uri '{\field{\*\fldinst{HYPERLINK "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace"}}{\fldrslt \AppleTypeServices\AppleTypeServicesF65539 \cf4 \ul \ulc4 \strokec4 https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace}}' -Method Get -Credential $creds -UseBasicParsing).content |\'a0ConvertFrom-Json\AppleTypeServices \'a0
\f1\fs24 \cb1 \

\f0\fs29\fsmilli14667 \AppleTypeServices\AppleTypeServicesF65539 \cb3 $messageTrace.d.results\'a0| Format-Table Index,\'a0MessageTraceId, Received,\'a0SenderAddress,\'a0FromIP,\'a0RecipientAddress,\'a0ToIP, Subject, Status, Size}
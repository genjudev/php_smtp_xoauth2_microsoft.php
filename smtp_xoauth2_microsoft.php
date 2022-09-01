<?php
/* composer.json
    "require": {
        "phpmailer/phpmailer": "^6.6",
        "league/oauth2-client": "^2.6",
        "thenetworg/oauth2-azure": "^2.1"
    }
*/
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\OAuth;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\PHPMailer;
use TheNetworg\OAuth2\Client\Provider\Azure;

require "vendor/autoload.php";

$mail = new PHPMailer(true);

$provider = new Azure([
    'clientId' => '',
    'clientSecret' => '',
    "scopes" => ["https://outlook.office.com/SMTP.Send"],
    "tenant" => "",
    "defaultEndPointVersion" => Azure::ENDPOINT_VERSION_2_0,
]);

$mail->setOAuth(
    new OAuth(
        [
            'provider' => $provider,
            'clientId' => '',
            'clientSecret' => '',
            'refreshToken' => '', 
            'userName' => 'mymail@office_365_email.tld',
        ]
    )
);

//Server settings
$mail->SMTPDebug = SMTP::DEBUG_SERVER;
$mail->isSMTP();
$mail->Host = 'smtp.office365.com';
$mail->Port = 587;
$mail->SMTPAuth = true;                                 
$mail->AuthType = 'XOAUTH2';
$mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;
$mail->CharSet = PHPMailer::CHARSET_UTF8;
//Recipients
$mail->setFrom('mymail@office_365_email.tld', 'name');
$mail->addAddress('spam@example.tld', 'Spam'); 

//Content
$mail->Subject = 'Here is the subject';
$mail->Body = 'Hallo';
$mail->AltBody = 'This is the body in plain text for non-HTML mail clients';

$mail->send();

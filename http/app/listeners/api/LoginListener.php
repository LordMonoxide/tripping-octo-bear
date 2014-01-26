<?php namespace api;

use Request;

use UserIP;

class LoginListener {
  public function onUserLogin($user) {
    $ip = $user->ips()->where('ip', '=', ip2long(Request::getClientIp()))->first();
    
    if($ip === null) {
      $ip = new UserIP;
      $ip->user()->associate($user);
      $ip->ip = ip2long(Request::getClientIp());
      $ip->authorised = true;
      $ip->save();
    } else {
      $ip->touch();
    }
  }
  
  public function subscribe($events) {
    $events->listen('auth.login', 'api\LoginListener@onUserLogin');
  }
}
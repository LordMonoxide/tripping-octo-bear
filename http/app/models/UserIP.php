<?php

class UserIP extends Eloquent {
  protected $table = 'user_ips';
  
  public function user() {
    return $this->belongsTo('User');
  }
}
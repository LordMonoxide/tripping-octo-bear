<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableUserIps extends Migration {
  public function up() {
    Schema::create('user_ips', function($table) {
      $table->increments('id');
      $table->integer('user_id')->unsigned();
      $table->string('ip', 15);
      $table->boolean('authorised')->default(false);
      $table->timestamps();
    });
  }
  
  public function down() {
    Schema::drop('user_ips');
  }
}
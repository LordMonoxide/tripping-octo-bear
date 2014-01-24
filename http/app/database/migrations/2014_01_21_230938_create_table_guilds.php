<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableGuilds extends Migration {
  public function up() {
    Schema::create('guilds', function($table) {
      $table->increments('id');
      $table->string('name', 40);
    });
  }
  
  public function down() {
    Schema::drop('guilds');
  }
}
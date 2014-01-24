<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableSkills extends Migration {
  public function up() {
    Schema::create('skills', function($table) {
      $table->increments('id');
      $table->string('name', 40);
      $table->string('desc', 256);
      $table->timestamps();
    });
  }
  
  public function down() {
    Schema::drop('skills');
  }
}
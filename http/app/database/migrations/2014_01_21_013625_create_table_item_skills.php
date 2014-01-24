<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableItemSkills extends Migration {
  public function up() {
    Schema::create('item_skills', function($table) {
      $table->increments('id');
      $table->enum('type', ['add', 'req']);
      $table->integer('item_id')->unsigned();
      $table->integer('skill_id')->unsigned();
      $table->integer('val');
      
      $table->foreign('item_id')
             ->references('id')
             ->on('items');
      
      $table->foreign('skill_id')
             ->references('id')
             ->on('skills');
    });
  }
  
  public function down() {
    Schema::drop('item_skills');
  }
}
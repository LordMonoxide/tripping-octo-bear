<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableCharacterSpells extends Migration {
  public function up() {
    Schema::create('character_spells', function($table) {
      $table->increments('id');
      $table->integer('character_id')->unsigned();
      $table->integer('spell_id')->unsigned();
      
      $table->foreign('character_id')
             ->references('id')
             ->on('characters');
      
      $table->foreign('spell_id')
             ->references('id')
             ->on('spells');
    });
  }
  
  public function down() {
    Schema::drop('character_spells');
  }
}
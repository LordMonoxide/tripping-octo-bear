<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableNpcSpells extends Migration {
  public function up() {
    Schema::create('npc_spells', function($table) {
      $table->increments('id');
      $table->integer('npc_id')->unsigned();
      $table->integer('spell_id')->unsigned();
      
      $table->foreign('npc_id')
             ->references('id')
             ->on('npcs');
      
      $table->foreign('spell_id')
             ->references('id')
             ->on('spells');
    });
  }
  
  public function down() {
    Schema::drop('npc_spells');
  }
}
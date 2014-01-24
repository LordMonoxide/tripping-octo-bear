<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableSpells extends Migration {
  public function up() {
    Schema::create('spells', function($table) {
      $table->increments('id');
      $table->string('name', 40);
      $table->string('desc', 256);
      $table->string('sound', 40);
      $table->enum('type', ['vital', 'warp', 'buff']);
      $table->integer('mpReq');
      $table->integer('levelReq');
      $table->integer('accessReq');
      $table->integer('castTime');
      $table->integer('cdTime');
      $table->integer('icon');
      $table->integer('map');
      $table->integer('x');
      $table->integer('y');
      $table->tinyInteger('dir');
      $table->integer('duration');
      $table->integer('interval');
      $table->tinyInteger('range');
      $table->boolean('isAOE');
      $table->integer('AOE');
      $table->integer('castAnim');
      $table->integer('spellAnim');
      $table->integer('stunDuration');
      $table->integer('hp');
      $table->integer('mp');
      $table->tinyInteger('hpType');
      $table->tinyInteger('mpType');
      $table->integer('buffType');
      $table->timestamps();
    });
  }
  
  public function down() {
    Schema::drop('spells');
  }
}
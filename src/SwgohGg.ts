// -tslint:disable: no-shadowed-variable
// -tslint:disable: object-literal-sort-keys

/** API Functions to pull data from swgoh.gg */
namespace SwgohGg {
  enum COMBAT_TYPE {
    HERO = 1,
    SHIP = 2,
  }

  interface AbilityData {
    id: string;
    is_zeta: boolean;
    name: string;
    ability_tier: number;
  }

  interface SwgohGgUnit {
    data: {
      base_id: string;
      combat_type: COMBAT_TYPE;
      gear: Array<{ base_id: string; is_obtained: boolean; slot: number }>;
      gear_level: number;
      level: number;
      power: number;
      rarity: number;
      ability_data: AbilityData[];
      stats: KeyedNumbers;
      url: string;
      zeta_abilities: string[];
    };
  }

  interface SwgohGgPlayerData {
    ally_code: number;
    arena: {
      members: string[];
      leader: string;
      rank: number;
    };
    fleet_arena: {
      members: string[];
      reinforcements: string[];
      leader: string;
      rank: number;
    };
    level: number;
    name: string;
    url: string;
    galactic_power: number;
    character_galactic_power: number;
    ship_galactic_power: number;
    ship_battles_won: number;
    pvp_battles_won: number;
    pve_battles_won: number;
    pve_hard_won: number;
    galactic_war_won: number;
    guild_raid_won: number;
    guild_contribution: number;
    guild_exchange_donations: number;
  }

  interface SwgohGgAbilityResponse {
    base_id: string;
    name: string;
    image: string;
    url: string;
    tier_max: number;
    is_zeta: boolean;
    is_omega: boolean;
    combat_type: number;
    type: number;
    character_base_id: string;
    ship_base_id: string;
  }

  interface SwgohGgUnitResponse {
    ability_classes: string[];
    alignment: string;
    base_id: string;
    categories: string[];
    combat_type: COMBAT_TYPE;
    description: string;
    gear_levels: Array<{ tier: number; gear: string[] }>;
    image: string;
    name: string;
    pk: number;
    power: number;
    role: string;
    url: string;
  }

  interface SwgohGgGuildResponse {
    data: {
      name: string;
      member_count: number;
      galactic_power: number;
      rank: number;
      profile_count: number;
      id: number;
    };
    players: SwgohGgPlayerResponse[];
  }

  interface SwgohGgPlayerResponse {
    data: SwgohGgPlayerData;
    units: SwgohGgUnit[];
  }

  /**
   * Send request to SwgohGg API
   * param link API 'GET' request
   * param errorMsg Message to display on error
   * returns JSON object response
   */
  function requestApi<T>(
    link: string,
    errorMsg: string = 'Error when retreiving data from swgoh.gg API',
  ): T | undefined {
    try {
      const response = UrlFetchApp.fetch(link, {
        // followRedirects: true,
        muteHttpExceptions: true,
      });
      return JSON.parse(response.getContentText()) as T;
    } catch (e) {
      // TODO: centralize alerts
      const UI = SpreadsheetApp.getUi();
      UI.alert(errorMsg, e, UI.ButtonSet.OK);
    }

    return undefined;
  }

  /**
   * Pull base Character data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getHeroList(): UnitDefinition[] {
    const json = requestApi<SwgohGgUnitResponse[]>('https://swgoh.gg/api/characters/');
    if (json) {
      const mapping = (unit: SwgohGgUnitResponse) => {
        return {
          abilities: [],
          alignment: unit.alignment.toLowerCase(),
          baseId: unit.base_id,
          name: unit.name,
          power: unit.power,
          role: unit.role.toLowerCase(),
          tags: unit.categories.map((category) => category.toLowerCase()),
          type: Units.TYPES.HERO,
        };
      };
      const units = json.map(mapping);
      const roles: string[] = [];
      for (const unit of units) {
        const role = unit.role;
        if (role !== 'leader' && roles.indexOf(role) === -1) {
          roles.push(role);
        }
      }
      for (const unit of units) {
        const unitRole = unit.role;
        if (unitRole === 'leader') {
          const tags = unit.tags;
          for (const role of roles) {
            if (tags.indexOf(role) > -1) {
              unit.role = role;
              break;
            }
          }
        }
      }

      return units;
    }

    return [];
  }

  /**
   * Pull base Ship data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getShipList(): UnitDefinition[] {
    const json = requestApi<SwgohGgUnitResponse[]>('https://swgoh.gg/api/ships/');
    if (json) {
      const mapping = (unit: SwgohGgUnitResponse) => {
        return {
          alignment: unit.alignment.toLowerCase(),
          baseId: unit.base_id,
          name: unit.name,
          power: unit.power,
          role: unit.role.toLowerCase(),
          tags: unit.categories.map((category) => category.toLowerCase()),
          type: Units.TYPES.SHIP,
        };
      };

      return json.map(mapping);
    }

    return [];
  }

  /**
   * Pull base Ability data from SwgohGg
   * returns Array of Abilites with [tags, baseId, name]
   */
  export function getAbilityList(): AbilityDefinition[] {
    const json = requestApi<SwgohGgAbilityResponse[]>('https://swgoh.gg/api/abilities/');
    if (json) {
      const mapping = (ability: SwgohGgAbilityResponse): AbilityDefinition => {
        return {
          baseId: ability.ship_base_id ? ability.ship_base_id : ability.character_base_id,
          isZeta: ability.is_zeta,
          name: ability.name,
          tierMax: ability.tier_max,
          type: ability.base_id.match(/^([^_]+)/)![1],
        };
      };

      return json.map(mapping);
    }

    return [];
  }

  /** Create guild API link */
  function getGuildApiLink(guildId: number): string {
    // TODO: data check
    return `https://swgoh.gg/api/guild/${guildId}/`;
  }

  /**
   * Pull Guild data from SwgohGg
   * Units name and tags are not populated
   * returns Array of Guild members and their units data
   */
  export function getGuildData(guildId: number): GuildData | undefined {
    const json = requestApi<SwgohGgGuildResponse>(getGuildApiLink(guildId));
    if (json && json.players) {
      const guild: GuildData = {
        id: guildId,
        members: [],
        name: json.data.name,
      };
      const members = guild.members;
      for (const member of json.players) {
        const unitArray: UnitInstances = {};
        for (const unit of member.units) {
          const data = unit.data;
          const type = data.combat_type === COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;
          const baseId = data.base_id;
          unitArray[baseId] = {
            abilities: data.ability_data.map(helper),
            baseId,
            gearLevel: data.gear_level,
            level: data.level,
            power: data.power,
            rarity: data.rarity,
            type,
          };
        }
        members.push({
          allyCode: +member.data.url.match(/(\d+)/)![1],
          fleetArenaBattlesWon: member.data.ship_battles_won,
          fleetArenaRank: member.data.fleet_arena ? member.data.fleet_arena.rank : undefined,
          galacticWarBattlesWon: member.data.galactic_war_won,
          gearDonatedInGuildExchange: member.data.guild_exchange_donations,
          gp: member.data.galactic_power,
          guildRaidsWon: member.data.guild_raid_won,
          guildTokensEarned: member.data.guild_contribution,
          hardBattlesWon: member.data.pve_hard_won,
          heroesGp: member.data.character_galactic_power,
          level: member.data.level,
          name: member.data.name,
          normalBattlesWon: member.data.pve_battles_won,
          shipsGp: member.data.ship_galactic_power,
          squadArenaBattlesWon: member.data.pvp_battles_won,
          squadArenaRank: member.data.arena ? member.data.arena.rank : undefined,
          units: unitArray,
        });
      }

      return guild;
    }

    return undefined;
  }

  /** Create player API link */
  function getPlayerApiLink(allyCode: number): string {
    // TODO: data check
    return `https://swgoh.gg/api/player/${allyCode}/`;
  }

  /**
   * Pull Player data from SwgohGg
   * Units name and tags are not populated
   * returns Player data, including its units data
   */
  export function getPlayerData(allyCode: number): PlayerData | undefined {
    const json = requestApi<SwgohGgPlayerResponse>(getPlayerApiLink(allyCode));
    if (json && json.data) {
      const data = json.data;
      const player: PlayerData = {
        allyCode: data.ally_code,
        fleetArenaBattlesWon: data.ship_battles_won,
        fleetArenaRank: data.fleet_arena ? data.fleet_arena.rank : undefined,
        galacticWarBattlesWon: data.galactic_war_won,
        gearDonatedInGuildExchange: data.guild_exchange_donations,
        gp: data.galactic_power,
        guildRaidsWon: data.guild_raid_won,
        guildTokensEarned: data.guild_contribution,
        hardBattlesWon: data.pve_hard_won,
        heroesGp: data.character_galactic_power,
        level: data.level,
        link: data.url,
        name: data.name,
        normalBattlesWon: data.pve_battles_won,
        shipsGp: data.ship_galactic_power,
        squadArenaBattlesWon: data.pvp_battles_won,
        squadArenaRank: data.arena ? data.arena.rank : undefined,
        units: {},
      };
      const units = player.units;
      for (const unit of json.units) {
        const unitData = unit.data;
        const type = unitData.combat_type === COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;
        const baseId = unitData.base_id;
        units[baseId] = {
          abilities: unitData.ability_data.map(helper),
          baseId,
          gearLevel: unitData.gear_level,
          level: unitData.level,
          power: unitData.power,
          rarity: unitData.rarity,
          type,
        };
      }

      return player;
    }

    return undefined;
  }

  const helper = (e: AbilityData) => {
    return {
      isZeta: e.is_zeta,
      name: e.name,
      tier: e.ability_tier,
      type: e.id.match(/^([^_]+)/)![1],
    };
  };
}

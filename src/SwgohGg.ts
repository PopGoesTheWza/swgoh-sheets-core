/** API Functions to pull data from swgoh.gg */
namespace SwgohGg {

  enum COMBAT_TYPE {
    HERO = 1,
    SHIP = 2,
  }

  interface SwgohGgUnit {
    data: {
      base_id: string;
      combat_type: COMBAT_TYPE;
      gear: {
        base_id: string;
        is_obtained: boolean;
        slot: number;
      }[];
      gear_level: number;
      level: number;
      power: number;
      rarity: number;
      ability_data: {
        id: string;
        is_zeta: boolean;
        name: string;
        ability_tier: number;
      }[];
      stats: KeyedNumbers;
      url: string;
      zeta_abilities: string[];
    };
  }

  interface SwgohGgPlayerData {
    ally_code: number;
    arena_leader_base_id: string;
    arena_rank: number;
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

  interface SwgohGgUnitResponse {
    ability_classes: string[];
    alignment: string;
    base_id: string;
    categories: string[];
    combat_type: COMBAT_TYPE;
    description: string;
    gear_levels: {
      tier: number;
      gear: string[];
    }[];
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
  ): T {

    let json;
    try {
      const params: URLFetchRequestOptions = {
        // followRedirects: true,
        muteHttpExceptions: true,
      };
      const response = UrlFetchApp.fetch(link, params);
      json = JSON.parse(response.getContentText());
    } catch (e) {
      // TODO: centralize alerts
      const UI = SpreadsheetApp.getUi();
      UI.alert(errorMsg, e, UI.ButtonSet.OK);
    }

    return json || undefined;
  }

  /**
   * Pull base Character data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getHeroList(): UnitDefinition[] {

    const json = requestApi<SwgohGgUnitResponse[]>('https://swgoh.gg/api/characters/');
    const mapping = (e: SwgohGgUnitResponse) => {
      const tags = [e.alignment, e.role, ...e.categories]
        .join(' ')  // TODO separator
        .toLowerCase();

      return { tags, baseId: e.base_id, name: e.name };
    };

    return json.map(mapping);
  }

  /**
   * Pull base Ship data from SwgohGg
   * returns Array of Characters with [tags, baseId, name]
   */
  export function getShipList(): UnitDefinition[] {

    const json = requestApi<SwgohGgUnitResponse[]>('https://swgoh.gg/api/ships/');
    const mapping = (e: SwgohGgUnitResponse) => {
      const tags = [e.alignment, e.role, ...e.categories]
        .join(' ')  // TODO separator
        .toLowerCase();

      return { tags, baseId: e.base_id, name: e.name };
    };

    return json.map(mapping);
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
  export function getGuildData(guildId: number): GuildData {

    const json = requestApi<SwgohGgGuildResponse>(getGuildApiLink(guildId));
    if (json && json.players) {
      const guild: GuildData = {
        id: guildId,
        name: json.data.name,
        members: [],
      };
      const members = guild.members;
      for (const member of json.players) {
        const unitArray: UnitInstances = {};
        for (const e of member.units) {
          const d = e.data;
          const type = d.combat_type === COMBAT_TYPE.HERO
            ? Units.TYPES.HERO
            : Units.TYPES.SHIP;
          const baseId = d.base_id;
          unitArray[baseId] = {
            type,
            baseId,
            gearLevel: d.gear_level,
            level: d.level,
            power: d.power,
            rarity: d.rarity,
            abilities: d.ability_data.map((e): Ability => {
              const type = e.id.match(/^([^_]+)/)[1];

              return { type, name: e.name, tier: e.ability_tier, isZeta: e.is_zeta };
            }),
          };
        }
        members.push({
          level: member.data.level,
          allyCode: +member.data.url.match(/(\d+)/)[1],
          name: member.data.name,
          gp: member.data.galactic_power,
          heroesGp: member.data.character_galactic_power,
          shipsGp: member.data.ship_galactic_power,
          fleetArenaBattlesWon: member.data.ship_battles_won,
          squadArenaBattlesWon: member.data.pvp_battles_won,
          normalBattlesWon: member.data.pve_battles_won,
          hardBattlesWon: member.data.pve_hard_won,
          galacticWarBattlesWon: member.data.galactic_war_won,
          guildRaidsWon: member.data.guild_raid_won,
          guildTokensEarned: member.data.guild_contribution,
          gearDonatedInGuildExchange: member.data.guild_exchange_donations,
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
  export function getPlayerData(allyCode: number): PlayerData {

    const json = requestApi<SwgohGgPlayerResponse>(getPlayerApiLink(allyCode));

    if (json && json.data) {
      const data = json.data;
      const player: PlayerData = {
        allyCode: data.ally_code,
        level: data.level,
        link: data.url,
        name: data.name,
        gp: data.galactic_power,
        heroesGp: data.character_galactic_power,
        shipsGp: data.ship_galactic_power,
        fleetArenaBattlesWon: data.ship_battles_won,
        squadArenaBattlesWon: data.pvp_battles_won,
        normalBattlesWon: data.pve_battles_won,
        hardBattlesWon: data.pve_hard_won,
        galacticWarBattlesWon: data.galactic_war_won,
        guildRaidsWon: data.guild_raid_won,
        guildTokensEarned: data.guild_contribution,
        gearDonatedInGuildExchange: data.guild_exchange_donations,
        units: {},
      };
      const units = player.units;
      for (const o of json.units) {
        const d = o.data;
        const type = d.combat_type === COMBAT_TYPE.HERO
          ? Units.TYPES.HERO
          : Units.TYPES.SHIP;
        const baseId = d.base_id;
        units[baseId] = {
          type,
          baseId,
          gearLevel: d.gear_level,
          level: d.level,
          power: d.power,
          rarity: d.rarity,
          abilities: d.ability_data.map((e): Ability => {
            const type = e.id.match(/^([^_]+)/)[1];

            return { type, name: e.name, tier: e.ability_tier, isZeta: e.is_zeta };
          }),
        };
      }

      return player;
    }

    return undefined;
  }

}

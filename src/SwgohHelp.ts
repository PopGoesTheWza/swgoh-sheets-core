import { swgohhelpapi } from '../lib';

/** API Functions to pull data from swgoh.help */
namespace SwgohHelp {

  /** Constants to translate categoryId into SwgohGg tags */
  enum categoryId {
    alignment_dark = 'dark side',
    alignment_light = 'light side',

    role_attacker = 'attacker',
    role_capital = 'capital ship',
    role_healer = 'healer',
    role_support = 'support',
    role_tank = 'tank',

    affiliation_empire = 'empire',
    affiliation_firstorder = 'first order',
    affiliation_imperialtrooper = 'imperial trooper',
    affiliation_nightsisters = 'nightsister',
    affiliation_oldrepublic = 'old republic',
    affiliation_phoenix = 'phoenix',
    affiliation_rebels = 'rebel',
    affiliation_republic = 'galactic republic',
    affiliation_resistance = 'resistance',
    affiliation_rogue_one = 'rogue one',
    affiliation_separatist = 'separatist',
    character_fleetcommander = 'fleet commander',
    profession_bountyhunter = 'bounty hunters',
    profession_clonetrooper = 'clone trooper',
    profession_jedi = 'jedi',
    profession_scoundrel = 'scoundrel',
    profession_sith = 'sith',
    profession_smuggler = 'smuggler',
    // shipclass_capitalship = 'capital ship',
    shipclass_cargoship = 'cargo ship',
    species_droid = 'droid',
    species_ewok = 'ewok',
    species_geonosian = 'geonosian',
    // species_human = 'human',
    species_jawa = 'jawa',
    species_tusken = 'tusken',
    // species_wookiee = 'wookiee',
  }

  /** check if swgohhelpapi api is installed */
  function checkLibrary(): boolean {

    const result = !!swgohhelpapi;
    if (!result) {
      const UI = SpreadsheetApp.getUi();
      UI.alert(
        'Library swgohhelpapi not found',
        `Please visit the link below to reinstall it.
https://github.com/PopGoesTheWza/swgoh-help-api/blob/master/README.md`,
        UI.ButtonSet.OK,
);
    }

    return result;
  }

  type UnitsList = {
    nameKey: string,
    forceAlignment: number,
    combatType: swgohhelpapi.COMBAT_TYPE,
    baseId: string,
    categoryIdList: [string],
  };

  /** Pull Units definitions from SwgohHelp */
  export function getUnitList(): UnitsDefinitions {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const units: [UnitsList] = client.fetchData({
      collection: swgohhelpapi.Collections.unitsList,
      language: swgohhelpapi.Languages.eng_us,
      match: {
        rarity: 7,
        obtainable: true,
        obtainableTime: 0,
      },
      project: {
        nameKey: true,
        forceAlignment: true,
        combatType: true,
        categoryIdList: true,
        baseId: true,
      },
    });

    if (units && units.length && units.length > 0) {
      return units.reduce(
        (acc: UnitsDefinitions, e) => {
          const bucket = e.combatType === swgohhelpapi.COMBAT_TYPE.HERO
            ? acc.heroes
            : acc.ships;
          const tags = e.categoryIdList.reduce(
            (a: [string], c) => {
              const tag = categoryId[c];
              if (tag) {
                a.push(tag);
              }
              return a;
            },
            [],
          );
          const alignment = e.forceAlignment === 2
            ? categoryId.alignment_light
            : (e.forceAlignment === 3 ? categoryId.alignment_dark : undefined);
          if (alignment) {
            tags.unshift(alignment);
          }
          const definition: UnitDefinition = {
            baseId: e.baseId,
            name: e.nameKey,
            tags: tags.unique().join(' '),  // TODO separator
          };
          bucket.push(definition);
          return acc;
        },
        { heroes: [], ships: [] });

    }

    return undefined;
  }

  /** Pull Guild data from SwgohHelp */
  export function getGuildData(allycode: number): GuildData {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings: swgohhelpapi.exports.Settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const jsonGuild: swgohhelpapi.exports.GuildResponse[] = client.fetchGuild({
      allycode,
      language: swgohhelpapi.Languages.eng_us,
      project: {
        id: true,
        name: true,
        members: true,
        // updated: true,
        roster: { allyCode: true, level: true, name: true },
      },
    });

    if (jsonGuild && jsonGuild.length && jsonGuild.length === 1) {

      const guild: GuildData = {
        id: +jsonGuild[0].id.match(/([\d]+)/)[1],
        name: jsonGuild[0].name,
        members: [],
      };

      const roster = jsonGuild[0].roster as swgohhelpapi.exports.PlayerResponse[];

      const members = roster.reduce(
        (acc: {allyCodes: number[]; membersData: PlayerData[]}, r) => {

          const allyCode = r.allyCode;
          if (allyCode) {
            const p: PlayerData = {
              allyCode,
              level: r.level,
              name: r.name,
              gp: undefined,
              heroesGp: undefined,
              shipsGp: undefined,
              fleetArenaRank: undefined,
              fleetArenaBattlesWon: undefined,
              squadArenaRank: undefined,
              squadArenaBattlesWon: undefined,
              normalBattlesWon: undefined,
              hardBattlesWon: undefined,
              galacticWarBattlesWon: undefined,
              guildRaidsWon: undefined,
              guildTokensEarned: undefined,
              gearDonatedInGuildExchange: undefined,
              units: {},
            };
            acc.allyCodes.push(allyCode);
            acc.membersData.push(p);
          }

          return acc;
        },
        {
          membersData: [],
          allyCodes: [],
        } as {allyCodes: number[]; membersData: PlayerData[]},
      );

      if (members.allyCodes.length > 0) {

        const jsonPlayers = client.fetchPlayer({
          allycodes: members.allyCodes,
          language: swgohhelpapi.Languages.eng_us,
          project: {
            allyCode: true,
            // guildName: true,
            level: true,
            name: true,
            roster: {
              combatType: true,
              defId: true,
              gp: true,
              gear: true,
              level: true,
              nameKey: true,
              rarity: true,
              // TODO: on demand
              skills: { id: true, tier: true, nameKey: true, isZeta: true },
            },
            stats: true,
            // updated: true,
          },
        });

        if (jsonPlayers && jsonPlayers.length && !jsonPlayers[0].hasOwnProperty('error')) {
          guild.members = jsonPlayers.map((e) => {

            const getStats = (i) => {
              const s = e.stats.find(o => o.index === i);
              return s && s.value;
            };
            const member: PlayerData = {
              allyCode: e.allyCode,
              level: e.level,
              name: e.name,
              gp: getStats(1),
              heroesGp: getStats(2),
              shipsGp: getStats(3),
              fleetArenaRank: e.arena.ship.rank,
              fleetArenaBattlesWon: getStats(4),
              squadArenaRank: e.arena.char.rank,
              squadArenaBattlesWon: getStats(5),
              normalBattlesWon: getStats(6),
              hardBattlesWon: getStats(7),
              galacticWarBattlesWon: getStats(8),
              guildRaidsWon: getStats(9),
              guildTokensEarned: getStats(10),
              gearDonatedInGuildExchange: getStats(11),
              units: {},
            };
            for (const u of e.roster) {
              const baseId = u.defId;
              const type = u.combatType === swgohhelpapi.COMBAT_TYPE.HERO
                ? Units.TYPES.HERO
                : Units.TYPES.SHIP;
              member.units[baseId] = {
                type,
                baseId: u.defId,
                gearLevel: u.gear,
                level: u.level,
                name: u.nameKey,
                power: u.gp,
                rarity: u.rarity,
                abilities: u.skills.map((e): Ability => {
                  const type = e.id.match(/^([^_]+)/)[1];
                  return { type, name: e.nameKey, tier: e.tier, isZeta: e.isZeta };
                }),
              };
            }

            return member;
          });

          return guild;
        }

      }
    }

    return undefined;
  }

  /**
   * Pull Player data from SwgohHelp
   * Units name and tags are not populated
   * returns Player data, including its units data
   */
  export function getPlayerData(allyCode: number): PlayerData {

    if (!checkLibrary()) {
      return undefined;
    }

    const settings = {
      username: config.SwgohHelp.username(),
      password: config.SwgohHelp.password(),
    };
    const client = new swgohhelpapi.exports.Client(settings);

    const json = client.fetchPlayer({
      allycodes: [allyCode],
      language: swgohhelpapi.Languages.eng_us,
      project: {
        allyCode: true,
        guildName: true,
        level: true,
        name: true,
        stats: true,
        roster: {
          combatType: true,
          defId: true,
          gp: true,
          gear: true,
          level: true,
          nameKey: true,
          rarity: true,
          // TODO: on demand
          skills: { id: true, tier: true, nameKey: true, isZeta: true },
        },
        arena: {
          char: {
            rank: true,
          },
          ship: {
            rank: true,
          },
        },
        // updated: true,
      },
    });

    if (json && json.length && json.length === 1 && !json[0].hasOwnProperty('error')) {
      const e = json[0];
      const getStats = (i) => {
        const s = e.stats.find(o => o.index === i);
        return s && s.value;
      };
      const player: PlayerData = {
        allyCode: e.allyCode,
        level: e.level,
        name: e.name,
        gp: getStats(1),
        heroesGp: getStats(2),
        shipsGp: getStats(3),
        fleetArenaRank: e.arena.ship.rank,
        fleetArenaBattlesWon: getStats(4),
        squadArenaRank: e.arena.char.rank,
        squadArenaBattlesWon: getStats(5),
        normalBattlesWon: getStats(6),
        hardBattlesWon: getStats(7),
        galacticWarBattlesWon: getStats(8),
        guildRaidsWon: getStats(9),
        guildTokensEarned: getStats(10),
        gearDonatedInGuildExchange: getStats(11),
        units: {},
      };
      for (const u of e.roster) {
        const baseId = u.defId;
        const type = u.combatType === swgohhelpapi.COMBAT_TYPE.HERO
          ? Units.TYPES.HERO
          : Units.TYPES.SHIP;
        player.units[baseId] = {
          type,
          baseId: u.defId,
          gearLevel: u.gear,
          level: u.level,
          name: u.nameKey,
          power: u.gp,
          rarity: u.rarity,
          abilities: u.skills.map((e): Ability => {
            const type = e.id.match(/^([^_]+)/)[1];
            return { type, name: e.nameKey, tier: e.tier, isZeta: e.isZeta };
          }),
        };
      }
      return player;
    }

    return undefined;
  }

}

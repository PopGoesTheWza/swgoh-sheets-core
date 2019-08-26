import { swgohhelpapi } from '../lib';

/** API Functions to pull data from swgoh.help */
namespace SwgohHelp {
  /** Constants to translate categoryId into SwgohGg tags */
  enum categoryId {
    alignment_dark = 'dark side',
    alignment_light = 'light side',
    alignment_neutral = 'neutral', // TODO check when actually used ingame

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

  interface UnitsList {
    nameKey: string;
    forceAlignment: number;
    combatType: swgohhelpapi.COMBAT_TYPE;
    baseId: string;
    categoryIdList: [string];
  }

  /** Pull Units definitions from SwgohHelp */
  export function getUnitList(): UnitsDefinitions | undefined {
    if (checkLibrary()) {
      const client = new swgohhelpapi.exports.Client({
        password: config.SwgohHelp.password(),
        username: config.SwgohHelp.username(),
      });

      const units: [UnitsList] = client.fetchData({
        collection: swgohhelpapi.Collections.unitsList,
        language: swgohhelpapi.Languages.eng_us,
        match: { obtainable: true, obtainableTime: 0, rarity: 7 },
        project: {
          baseId: true,
          categoryIdList: true,
          combatType: true,
          crewList: false,
          descKey: false,
          forceAlignment: true,
          nameKey: true,
          skillReferenceList: false,
        },
      });

      if (units && units.length && units.length > 0) {
        return units.reduce(
          (acc: UnitsDefinitions, unitDef) => {
            const bucket = unitDef.combatType === swgohhelpapi.COMBAT_TYPE.HERO ? acc.heroes : acc.ships;
            const tags = unitDef.categoryIdList.reduce(
              (accTags: { alignment?: string; role?: string; tags: string[] }, category) => {
                const tag = categoryId[category as keyof typeof categoryId];
                if (tag) {
                  if (category.indexOf('alignment_') > -1) {
                    accTags.alignment = tag;
                  } else if (category.indexOf('role_') > -1) {
                    accTags.role = tag;
                  } else {
                    accTags.tags.push(tag);
                  }
                }
                return accTags;
              },
              { alignment: undefined, role: undefined, tags: [] as string[] },
            );
            tags.alignment =
              unitDef.forceAlignment === 2
                ? categoryId.alignment_light
                : unitDef.forceAlignment === 3
                ? categoryId.alignment_dark
                : categoryId.alignment_neutral;
            const definition: UnitDefinition = {
              alignment: tags.alignment,
              baseId: unitDef.baseId,
              name: unitDef.nameKey,
              power: 0,
              role: tags.role!,
              tags: tags.tags,
            };
            bucket.push(definition);
            return acc;
          },
          { heroes: [], ships: [] },
        );
      }
    }

    return undefined;
  }

  /** Pull Guild data from SwgohHelp */
  export function getGuildData(allycode: number): GuildData | undefined {
    if (checkLibrary()) {
      const client = new swgohhelpapi.exports.Client({
        password: config.SwgohHelp.password(),
        username: config.SwgohHelp.username(),
      });
      const jsonGuild: swgohhelpapi.exports.GuildResponse[] = client.fetchGuild({
        allycode,
        enums: true,
        language: swgohhelpapi.Languages.eng_us,
        project: {
          gp: false,
          members: true,
          name: true,
          roster: {
            allyCode: true,
            gp: false,
            gpChar: false,
            gpShip: false,
            guildMemberLevel: true,
            level: true,
            name: true,
          },
          // updated: true,
        },
      });

      if (jsonGuild && jsonGuild.length && jsonGuild.length === 1) {
        const guild: GuildData = {
          id: +jsonGuild[0].id.match(/([\d]+)/)![1],
          members: [],
          name: jsonGuild[0].name!,
        };

        const roster = jsonGuild[0].roster as swgohhelpapi.exports.PlayerResponse[];

        const members = roster.reduce(
          (acc: { allyCodes: number[]; membersData: PlayerData[] }, member) => {
            const allyCode = member.allyCode;
            if (allyCode) {
              const p: PlayerData = {
                allyCode,
                fleetArenaBattlesWon: undefined,
                fleetArenaRank: undefined,
                galacticWarBattlesWon: undefined,
                gearDonatedInGuildExchange: undefined,
                gp: 0,
                guildRaidsWon: undefined,
                guildTokensEarned: undefined,
                hardBattlesWon: undefined,
                heroesGp: 0,
                level: member.level!,
                name: member.name!,
                normalBattlesWon: undefined,
                shipsGp: 0,
                squadArenaBattlesWon: undefined,
                squadArenaRank: undefined,
                units: {},
              };
              acc.allyCodes.push(allyCode);
              acc.membersData.push(p);
            }

            return acc;
          },
          { allyCodes: [], membersData: [] },
        );

        if (members.allyCodes.length > 0) {
          const jsonPlayers = client.fetchPlayer({
            allycodes: members.allyCodes,
            enums: true,
            language: swgohhelpapi.Languages.eng_us,
            project: {
              allyCode: true,
              arena: { char: { rank: true }, ship: { rank: true } },
              // grandArena: false,
              // grandArenaLifeTime: false,
              // guildName: true,
              level: true,
              name: true,
              roster: {
                combatType: true,
                crew: false,
                defId: true,
                equipped: false,
                gear: true,
                gp: true,
                level: true,
                mods: false,
                nameKey: true,
                rarity: true,
                // primaryUnitStat: false,
                // TODO: on demand
                skills: { id: true, isZeta: true, nameKey: true, tier: true },
              },
              stats: true,
              // updated: true,
            },
          });

          if (jsonPlayers && jsonPlayers.length && !jsonPlayers[0].hasOwnProperty('error')) {
            guild.members = jsonPlayers.map((player) => {
              const getStats = (i: number) => {
                const s = player.stats!.find((o) => o.index === i);
                return s && s.value;
              };
              const member: PlayerData = {
                allyCode: player.allyCode!,
                fleetArenaBattlesWon: getStats(4),
                fleetArenaRank: player.arena!.ship.rank,
                galacticWarBattlesWon: getStats(8),
                gearDonatedInGuildExchange: getStats(11),
                gp: getStats(1)!,
                guildRaidsWon: getStats(9),
                guildTokensEarned: getStats(10),
                hardBattlesWon: getStats(7),
                heroesGp: getStats(2)!,
                level: player.level!,
                name: player.name!,
                normalBattlesWon: getStats(6),
                shipsGp: getStats(3)!,
                squadArenaBattlesWon: getStats(5),
                squadArenaRank: player.arena!.char.rank,
                units: {},
              };
              for (const unit of player.roster!) {
                const baseId = unit.defId;
                const type = unit.combatType === swgohhelpapi.COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;
                member.units[baseId] = {
                  abilities: unit.skills.map(
                    (e): Ability => ({
                      isZeta: e.isZeta,
                      name: e.nameKey,
                      tier: e.tier,
                      type: e.id.match(/^([^_]+)/)![1],
                    }),
                  ),
                  baseId: unit.defId,
                  gearLevel: unit.gear,
                  level: unit.level,
                  name: unit.nameKey,
                  power: unit.gp,
                  rarity: unit.rarity,
                  type,
                };
              }

              return member;
            });

            return guild;
          }
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
  export function getPlayerData(allyCode: number): PlayerData | undefined {
    if (!checkLibrary()) {
      return undefined;
    }

    const client = new swgohhelpapi.exports.Client({
      password: config.SwgohHelp.password(),
      username: config.SwgohHelp.username(),
    });

    const json = client.fetchPlayer({
      allycodes: [allyCode],
      enums: true,
      language: swgohhelpapi.Languages.eng_us,
      project: {
        allyCode: true,
        arena: { char: { rank: true }, ship: { rank: true } },
        // grandArena: false,
        // grandArenaLifeTime: false,
        guildName: true,
        level: true,
        name: true,
        roster: {
          combatType: true,
          crew: false,
          defId: true,
          equipped: false,
          gear: true,
          gp: true,
          level: true,
          mods: false,
          nameKey: true,
          rarity: true,
          // primaryUnitStat: false,
          // TODO: on demand
          skills: { id: true, isZeta: true, nameKey: true, tier: true },
          xp: false,
        },
        stats: true,
        // titles: false,
      },
    });

    if (json && json.length && json.length === 1 && !json[0].hasOwnProperty('error')) {
      const playerResponse = json[0];
      const getStats = (i: number) => {
        const s = playerResponse.stats!.find((o) => o.index === i);
        return s && s.value;
      };
      const player: PlayerData = {
        allyCode: playerResponse.allyCode!,
        fleetArenaBattlesWon: getStats(4),
        fleetArenaRank: playerResponse.arena!.ship.rank,
        galacticWarBattlesWon: getStats(8),
        gearDonatedInGuildExchange: getStats(11),
        gp: getStats(1)!,
        guildRaidsWon: getStats(9),
        guildTokensEarned: getStats(10),
        hardBattlesWon: getStats(7),
        heroesGp: getStats(2)!,
        level: playerResponse.level!,
        name: playerResponse.name!,
        normalBattlesWon: getStats(6),
        shipsGp: getStats(3)!,
        squadArenaBattlesWon: getStats(5),
        squadArenaRank: playerResponse.arena!.char.rank,
        units: {},
      };
      for (const u of playerResponse.roster!) {
        const baseId = u.defId;
        const type = u.combatType === swgohhelpapi.COMBAT_TYPE.HERO ? Units.TYPES.HERO : Units.TYPES.SHIP;
        player.units[baseId] = {
          abilities: u.skills.map(
            (e): Ability => ({ isZeta: e.isZeta, name: e.nameKey, tier: e.tier, type: e.id.match(/^([^_]+)/)![1] }),
          ),
          baseId: u.defId,
          gearLevel: u.gear,
          level: u.level,
          name: u.nameKey,
          power: u.gp,
          rarity: u.rarity,
          type,
        };
      }
      return player;
    }

    return undefined;
  }
}

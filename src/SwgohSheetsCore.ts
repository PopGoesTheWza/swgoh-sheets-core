/** refresh guilds data */
function refreshData(): void {
  Core.refreshData();
}

/** workaround to tslint issue of namespace scope after importingtype definitions */
declare namespace SwgohHelp {
  function getGuildData(allycode: number): GuildData;
  function getPlayerData(allyCode: number): PlayerData;
  function getUnitList(): UnitsDefinitions;
}

/** Shortcuts for Google Apps Script classes */
const SPREADSHEET = SpreadsheetApp.getActive();

import Spreadsheet = GoogleAppsScript.Spreadsheet;
import URL_Fetch = GoogleAppsScript.URL_Fetch;
// type URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

interface KeyedType<T> {
  [key: string]: T;
}

type KeyedNumbers = KeyedType<number>;

interface GuildData {
  id: number;
  name: string;
  members: PlayerData[];
}

interface PlayerData {
  allyCode: number;
  level: number;
  link?: string;
  name: string;
  gp: number;
  heroesGp: number;
  shipsGp: number;
  fleetArenaRank?: number;
  fleetArenaBattlesWon?: number;
  squadArenaRank?: number;
  squadArenaBattlesWon?: number;
  normalBattlesWon?: number;
  hardBattlesWon?: number;
  galacticWarBattlesWon?: number;
  guildRaidsWon?: number;
  guildTokensEarned?: number;
  gearDonatedInGuildExchange?: number;
  units: UnitInstances;
}

interface UnitsDefinitions {
  heroes: UnitDefinition[];
  ships: UnitDefinition[];
}

/** A unit's name, baseId and tags */
interface UnitDefinition {
  /** Unit Id */
  baseId: string;
  name: string;
  /** Alignment, role and tags */
  alignment: string;
  role: string;
  tags: string[];
  type?: number;
  abilities?: Ability[];
  power: number;
}

interface AbilityDefinition {
  baseId: string;
  name: string;
  type: string;
  tierMax: number;
  isZeta: boolean;
}

interface Ability {
  name: string;
  type: string;
  tier: number;
  isZeta: boolean;
}

/**
 * A unit instance attributes
 * (baseId, gearLevel, level, name, power, rarity, stats, tags)
 */
interface UnitInstance {
  type: Units.TYPES;
  baseId?: string;
  gearLevel?: number;
  level: number;
  name?: string;
  power: number;
  abilities: Ability[];
  rarity: number;
  stats?: string;
  tags?: string;
}

interface UnitInstances {
  [key: string]: UnitInstance;
}

/** Constants for data source */
enum DATASOURCES {
  /** Use swgoh.gg API as data source */
  SWGOH_GG = 'SwgohGg',
  /** Use swgoh.help API as data source */
  SWGOH_HELP = 'SwgohHelp',
}

/** Constants for sheets name */
enum CORESHEET {
  // SETUP = 'coreSetup',
  ROSTER = 'coreRoster',
  UNITS = 'coreUnits',
  ABILITIES = 'coreAbilities',
  HEROES = 'coreHeroes',
  SHIPS = 'coreShips',
}

/** Constants for sheets name */
enum CORERANGE {
  ADDMEMBERGPTOROSTER = 'AddMemberGPToRoster',
  ADDMEMBERBATTLESTOROSTER = 'AddMemberBattlesToRoster',
  ADDMEMBERGUILDACTIVITIESTOROSTER = 'AddMemberGuildActivitiesToRoster',
  WHICHUNITTYPE = 'WhichUnitType',
  WHICHHEROABILITIES = 'WhichHeroAbilities',
  GETSHIPABILITIES = 'GetShipAbilities',
  SWGOHHELPUSERNAME = 'SwgohHelpUsername',
  SWGOHHELPPASSWORD = 'SwgohHelpPassword',
  GUILDDATASOURCESETTINGS = 'GuildDataSourceSettings',
  RENAMEADDREMOVE = 'RenameAddRemove',
}

/** settings related functions */
namespace config {
  function getNemedRangeAsString(name: Spreadsheet.Range | string) {
    return (typeof name === 'string' ? SPREADSHEET.getRangeByName(name) : name).getValue() as string;
  }

  /** use members galactic power stats */
  export function addMemberGPToRoster() {
    return getNemedRangeAsString(CORERANGE.ADDMEMBERGPTOROSTER) === 'ON';
  }

  /** use members galactic power stats */
  export function addMemberBattlesToRoster() {
    return getNemedRangeAsString(CORERANGE.ADDMEMBERBATTLESTOROSTER) === 'ON';
  }

  /** use members galactic power stats */
  export function addMembersGuildActivitiesToRoster() {
    return getNemedRangeAsString(CORERANGE.ADDMEMBERGUILDACTIVITIESTOROSTER) === 'ON';
  }

  export enum UNIT_TYPES {
    BOTH = 'Heroes & Ships',
    HERO = 'Heroes only',
    SHIP = 'Ships only',
  }

  /** use members galactic power stats */
  export function whichUnitType() {
    return getNemedRangeAsString(CORERANGE.WHICHUNITTYPE) as UNIT_TYPES;
  }

  export enum HERO_ABILITIES {
    NONE = 'None',
    ZETAS = 'Zetas only',
    WORTHY = 'Leader & zetas',
    DETAILED = 'Detailed',
  }

  /** use members galactic power stats */
  export function whichHeroAbilities() {
    return getNemedRangeAsString(CORERANGE.WHICHHEROABILITIES) as HERO_ABILITIES;
  }

  /** use members galactic power stats */
  export function getShipAbilities() {
    return getNemedRangeAsString(CORERANGE.GETSHIPABILITIES) === 'ON';
  }

  /** SwgohHelp related settings */
  // tslint:disable-next-line: no-shadowed-variable
  export namespace SwgohHelp {
    /** Get the SwgohHelp API username */
    export function username() {
      return `${getNemedRangeAsString(CORERANGE.SWGOHHELPUSERNAME)}`.trim();
    }

    /** Get the SwgohHelp API password */
    export function password(): string {
      return `${getNemedRangeAsString(CORERANGE.SWGOHHELPPASSWORD)}`;
    }
  }
}

namespace Units {
  export enum TYPES {
    HERO = 1,
    SHIP = 2,
  }
}

namespace Core {
  enum UnitsAbilities {
    NONE = 'None',
    ZETAS_ONLY = 'Zetas only',
    ALL = 'All',
  }

  // tslint:disable-next-line: no-shadowed-variable
  export const refreshData = (): void => {
    const activeGuilds = getActiveGuilds();
    const credentialIsFresh = isCredentialFresh();
    const rarIsFresh = isRARFresh();
    const guildsIsFresh = isGuildsFresh();

    let staleGuilds = guildsIsFresh && credentialIsFresh && rarIsFresh ? undefined : activeGuilds;

    // let staleGuilds = activeGuilds.filter((e) => {

    //   if (rarIsFresh) {

    //     if (e.dataSource === DATASOURCES.SWGOH_GG) {
    //       return !isGuildsFresh();
    //     }
    //     if (e.dataSource === DATASOURCES.SWGOH_HELP) {
    //       return !(credentialIsFresh && isGuildFresh(e));
    //     }

    //   } else {  // TODO: intersect rar and guilds
    //     return true;
    //   }

    // });

    if (staleGuilds) {
      staleGuilds = activeGuilds; // TODO proper handling of stale

      // SPREADSHEET.toast('Get base unit definitions', 'Refresh data', 5);
      const baseUnits = getBaseUnits();
      const datas = staleGuilds.map((e) => getGuildData(e)) as GuildData[];
      renameAddRemove(staleGuilds, datas);
      setGuildNames(staleGuilds, datas);
      writeGuildNames(activeGuilds);
      // SPREADSHEET.toast(`Writing ${CORESHEETS.ROSTER}`, 'Refresh data', 5);
      writeRoster(datas);
      // SPREADSHEET.toast(`Writing ${CORESHEETS.UNITS}`, 'Refresh data', 5);
      writeBaseUnits(baseUnits);
      // SPREADSHEET.toast(`Writing ${CORESHEETS.ABILITIES}`, 'Refresh data', 5);
      writeBaseAbilities(getBaseAbilities());
      writeUnits(datas, baseUnits);
      refreshGuilds();
      refreshCredential();
      refreshRAR();
    } else {
      SPREADSHEET.toast('Data from API is cached for one hour', 'Nothing to refresh', 5);

      return;
    }
  };

  const renameAddRemove = (settings: GuildSettings[], datas: GuildData[]): void => {
    const rar = getRAR(datas);
    for (const add of rar.add) {
      const { allyCode, guild } = add;
      const isInGuild = datas.find((e) => !!e.members.find((m) => m.allyCode === allyCode));
      const targetGuild = datas.find((e) => e.name === guild);
      if (targetGuild) {
        // is in another guild?
        if (isInGuild) {
          if (isInGuild !== targetGuild) {
            targetGuild.members.push(
              isInGuild.members.splice(isInGuild.members.findIndex((e) => e.allyCode === allyCode), 1)[0],
            );
          }
        } else {
          const useDotHelp = `${config.SwgohHelp.password()}`.trim().length > 0;
          // try to guess data source
          // SPREADSHEET.toast(
          //   `Get data for member ${allyCode} of guild ${targetGuild.id} from ${useDotHelp ? 'swgoh.help' : 'swgoh.gg'}`,
          //   'Refresh data',
          //   5,
          // );
          const player = (useDotHelp ? SwgohHelp : SwgohGg).getPlayerData(allyCode);
          if (player) {
            targetGuild.members.push(player);
          }
        }
      } else if (guild === 'PLAYER') {
        const target: { id: number; name: string; members: PlayerData[] } = { id: 0, name: guild, members: [] };
        datas.push(target);
        const useDotHelp = `${config.SwgohHelp.password()}`.trim().length > 0;
        // SPREADSHEET.toast(
        //   `Get data for player ${allyCode} from ${useDotHelp ? 'swgoh.help' : 'swgoh.gg'}`,
        //   'Refresh data',
        //   5,
        // );
        const player = (useDotHelp ? SwgohHelp : SwgohGg).getPlayerData(allyCode);
        if (player) {
          target.members.push(player);
        }
      }
    }
    for (const rename of rar.rename) {
      const allyCode = rename.allyCode;
      for (const data of datas) {
        const i = data.members.findIndex((member) => member.allyCode === allyCode);
        if (i > -1) {
          data.members[i].name = rename.name;
        }
      }
    }
    for (const allyCode of rar.remove) {
      for (const data of datas) {
        const i = data.members.findIndex((member) => member.allyCode === allyCode);
        if (i > -1) {
          data.members.splice(i, 1);
        }
      }
    }
  };

  const setGuildNames = (settings: GuildSettings[], datas: GuildData[]): void => {
    // tslint:disable-next-line: forin
    for (const i in settings) {
      const setup = settings[i];
      if (setup.name.length === 0) {
        setup.name = datas[i].name;
      }
    }
    // TODO: duplicates guild name
  };

  const writeGuildNames = (settings: GuildSettings[]): void => {
    const range = SPREADSHEET.getRangeByName(CORERANGE.GUILDDATASOURCESETTINGS);
    const values = range.getValues() as string[][];
    for (const setting of settings) {
      const name = setting.name;
      const dataSource = setting.dataSource;
      const row = values.find(
        (e) =>
          e[0] === 'ON' &&
          e[2] === dataSource &&
          (dataSource === DATASOURCES.SWGOH_GG ? +e[3] === setting.guildId : +e[4] === setting.allyCode),
      );
      if (row && row[1] !== name) {
        row[1] = name;
      }
    }
    range.setValues(values);
  };

  const writeRoster = (datas: GuildData[]): void => {
    const addGP = config.addMemberGPToRoster();
    const addBattles = config.addMemberBattlesToRoster();
    const addActivities = config.addMembersGuildActivitiesToRoster();
    const values: any[][] = [];
    for (const data of datas) {
      if (data) {
        const guildId = data.id;
        const guildName = data.name;
        for (const member of data.members) {
          let columns = [guildId, guildName, member.name, member.allyCode, member.level];
          if (addGP) {
            columns = [...columns, ...[member.gp, member.heroesGp, member.shipsGp]];
          }
          if (addBattles) {
            columns = [
              ...columns,
              ...[
                member.fleetArenaRank!,
                member.fleetArenaBattlesWon!,
                member.squadArenaRank!,
                member.squadArenaBattlesWon!,
                member.normalBattlesWon!,
                member.hardBattlesWon!,
                member.galacticWarBattlesWon!,
              ],
            ];
          }
          if (addActivities) {
            columns = [
              ...columns,
              ...[member.guildRaidsWon!, member.guildTokensEarned!, member.gearDonatedInGuildExchange!],
            ];
          }
          values.push(columns);
        }
      }
    }
    let headers = ['guildId', 'guildName', 'name', 'allyCode', 'level'];

    if (addGP) {
      headers = [...headers, ...['gp', 'heroesGp', 'shipsGp']];
    }
    if (addBattles) {
      headers = [
        ...headers,
        ...[
          'fleetArenaRank',
          'fleetArenaBattlesWon',
          'squadArenaRank',
          'squadArenaBattlesWon',
          'normalBattlesWon',
          'hardBattlesWon',
          'galacticWarBattlesWon',
        ],
      ];
    }
    if (addActivities) {
      headers = [...headers, ...['guildRaidsWon', 'guildTokensEarned', 'gearDonatedInGuildExchange']];
    }
    Sheets.setValues(Sheets.getSheetByNameOrDie(CORESHEET.ROSTER), values, headers);
  };

  const abilityTypeOrder = [
    'hardwareskill',
    'contractskill',
    'uniqueskill',
    'leaderskill',
    'specialskill',
    'basicskill',
  ];

  const sortAbilities = (a: Ability, b: Ability): number => {
    const types = abilityTypeOrder.indexOf(b.type) - abilityTypeOrder.indexOf(a.type);
    return types === 0 ? a.name.localeCompare(b.name) : types;
  };

  const writeBaseUnits = (baseUnits: UnitDefinition[]): void => {
    Sheets.setValues(
      Sheets.getSheetByNameOrDie(CORESHEET.UNITS),
      baseUnits.map((e) => [e.name, e.baseId, e.type, e.alignment, e.role, e.tags.join(','), e.power]),
      ['name', 'baseId', 'type', 'alignment', 'role', 'tags', 'power'],
    );
  };

  const writeBaseAbilities = (baseAbilities: AbilityDefinition[]): void => {
    Sheets.setValues(
      Sheets.getSheetByNameOrDie(CORESHEET.ABILITIES),
      baseAbilities.map((e) => [e.baseId, e.name, e.type, e.tierMax, e.isZeta]),
      ['baseId', 'name', 'type', 'tierMax', 'isZeta'],
    );
  };

  const writeUnits = (datas: GuildData[], baseUnits: UnitDefinition[]): void => {
    // TODO: use distinct sheet for base unit data (tags, abilities...)
    const unitTypes = config.whichUnitType();
    const addHeroes = unitTypes !== config.UNIT_TYPES.SHIP;
    const addShips = unitTypes !== config.UNIT_TYPES.HERO;
    const heroAbilities = config.whichHeroAbilities();
    const getAblilities = heroAbilities !== config.HERO_ABILITIES.NONE;
    const getZetaAblilities = heroAbilities === config.HERO_ABILITIES.ZETAS;
    const getWorthyAblilities = heroAbilities === config.HERO_ABILITIES.WORTHY;
    const getAllAblilities = heroAbilities === config.HERO_ABILITIES.DETAILED;
    const shipAbilities = config.getShipAbilities();
    const heroes: any[][] = [];
    const ships: any[][] = [];
    const heroesAbilities = { count: 0, columns: 0 };
    const shipsAbilities = { count: 0, columns: 0 };
    for (const data of datas) {
      if (data) {
        for (const member of data.members) {
          const units = member.units;
          // tslint:disable-next-line: forin
          for (const baseId in units) {
            const def = baseUnits.find((e) => e.baseId === baseId) as UnitDefinition;
            const unit = units[baseId];
            // sort abilities
            unit.abilities.sort(sortAbilities);
            // unit.abilities = unit.abilities.sort(sortAbilities);
            if (!def.abilities) {
              def.abilities = unit.abilities;
            }
            const type = def.type || unit.type;
            if (!def.type) {
              def.type = type;
            }
            if ((addHeroes && type === Units.TYPES.HERO) || (addShips && type === Units.TYPES.SHIP)) {
              const columns =
                type === Units.TYPES.HERO
                  ? [
                      member.name,
                      member.allyCode,
                      def.name,
                      // def.tags,
                      unit.rarity,
                      unit.level,
                      unit.gearLevel,
                      unit.power,
                    ]
                  : [
                      member.name,
                      member.allyCode,
                      def.name,
                      // def.tags,
                      unit.rarity,
                      unit.level,
                      unit.power,
                    ];
              if (getAblilities && type === Units.TYPES.HERO) {
                if (getZetaAblilities) {
                  let count = 0;
                  for (const ability of unit.abilities) {
                    if (ability.isZeta && ability.tier === 8) {
                      columns.push(ability.name);
                      columns.push(ability.type);
                      count += 1;
                    }
                  }
                  heroesAbilities.count = Math.max(count, heroesAbilities.count);
                } else if (getWorthyAblilities) {
                  let count = 0;
                  for (const ability of unit.abilities) {
                    if (ability.type === 'leaderskill' || (ability.isZeta && ability.tier === 8)) {
                      columns.push(ability.name);
                      columns.push(ability.type);
                      columns.push(ability.isZeta ? 'true' : 'false');
                      columns.push(ability.tier);
                      count += 1;
                    }
                  }
                  heroesAbilities.count = Math.max(count, heroesAbilities.count);
                } else if (getAllAblilities) {
                  for (const ability of unit.abilities) {
                    columns.push(ability.name);
                    columns.push(ability.type);
                    columns.push(ability.isZeta ? 'true' : 'false');
                    columns.push(ability.tier);
                  }
                  heroesAbilities.count = Math.max(unit.abilities.length, heroesAbilities.count);
                }
                heroesAbilities.columns = Math.max(columns.length, heroesAbilities.columns);
              } else if (shipAbilities && type === Units.TYPES.SHIP) {
                // TODO: write abilites (enum UnitsAbilities)
                for (const ability of unit.abilities) {
                  columns.push(ability.name);
                  columns.push(ability.type);
                  columns.push(ability.tier);
                }
                shipsAbilities.count = Math.max(unit.abilities.length, shipsAbilities.count);
                shipsAbilities.columns = Math.max(columns.length, shipsAbilities.columns);
              }
              (type === Units.TYPES.HERO ? heroes : ships).push(columns);
            }
          }
        }
      }
    }
    const heroesHeaders = [
      'name',
      'allyCode',
      'unit',
      // 'tags',
      'rarity',
      'level',
      'gearLevel',
      'power',
    ];
    const shipsHeaders = [
      'name',
      'allyCode',
      'unit',
      // 'tags',
      'rarity',
      'level',
      'power',
    ];
    if (getZetaAblilities) {
      for (let i = 0; i < heroesAbilities.count; i += 1) {
        heroesHeaders.push(`ability_${i}`);
        heroesHeaders.push(`type_${i}`);
      }
    } else if (getWorthyAblilities || getAllAblilities) {
      for (let i = 0; i < heroesAbilities.count; i += 1) {
        heroesHeaders.push(`ability_${i}`);
        heroesHeaders.push(`type_${i}`);
        heroesHeaders.push(`isZeta_${i}`);
        heroesHeaders.push(`tier_${i}`);
      }
    }
    if (shipAbilities) {
      for (let i = 0; i < shipsAbilities.count; i += 1) {
        shipsHeaders.push(`ability_${i}`);
        shipsHeaders.push(`type_${i}`);
        shipsHeaders.push(`tier_${i}`);
      }
    }
    const counts = baseUnits.reduce(
      (acc, e) => {
        // acc.tags = Math.max(acc.tags, e.tags.length);
        acc.abilities = Math.max(acc.abilities, e.abilities ? e.abilities.length : 0);
        return acc;
      },
      { tags: 0, abilities: 0 },
    );
    const unitsHeaders = ['name', 'baseId', 'type', 'alignment', 'role', 'tags'];
    if (counts.tags) {
      for (let i = 0; i < counts.tags; i += 1) {
        unitsHeaders.push(`tag_${i}`);
      }
    }
    if (counts.abilities) {
      for (let i = 0; i < counts.abilities; i += 1) {
        unitsHeaders.push(`ability_${i}`);
        unitsHeaders.push(`type_${i}`);
        unitsHeaders.push(`isZeta_${i}`);
      }
    }
    heroesAbilities.columns = Math.max(heroesHeaders.length, heroesAbilities.columns);
    // SPREADSHEET.toast(`Writing ${CORESHEETS.HEROES}`, 'Refresh data', 5);
    Sheets.setValues(
      Sheets.getSheetByNameOrDie(CORESHEET.HEROES),
      heroes.map((e) =>
        e.length !== heroesAbilities.columns
          ? [...e, ...Array(heroesAbilities.columns).fill(null)].slice(0, heroesAbilities.columns)
          : e,
      ),
      heroesHeaders,
    );
    shipsAbilities.columns = Math.max(shipsHeaders.length, shipsAbilities.columns);
    // SPREADSHEET.toast(`Writing ${CORESHEETS.SHIPS}`, 'Refresh data', 5);
    Sheets.setValues(
      Sheets.getSheetByNameOrDie(CORESHEET.SHIPS),
      ships.map((e) =>
        e.length !== shipsAbilities.columns
          ? [...e, ...Array(shipsAbilities.columns).fill(null)].slice(0, shipsAbilities.columns)
          : e,
      ),
      shipsHeaders,
    );
  };

  interface GuildSettings {
    name: string;
    dataSource: DATASOURCES;
    guildId?: number;
    allyCode?: number;
  }

  /** returns an array of active guild settings */
  const getActiveGuilds = (): GuildSettings[] => {
    const guilds = SPREADSHEET.getRangeByName(CORERANGE.GUILDDATASOURCESETTINGS)
      .getValues()
      .reduce((acc: GuildSettings[], columns: string[]) => {
        if (columns[0] === 'ON') {
          // hardcoded checkbox value
          const name = columns[1];
          const dataSource = columns[2] as DATASOURCES;
          if (dataSource === DATASOURCES.SWGOH_GG) {
            acc.push({ name, dataSource, guildId: +columns[3] });
          } else if (dataSource === DATASOURCES.SWGOH_HELP) {
            acc.push({ name, dataSource, allyCode: +columns[4] });
          }
        }

        return acc;
      }, []) as GuildSettings[];

    return guilds;
  };

  const getBaseUnits = (): UnitDefinition[] => {
    return [...SwgohGg.getHeroList(), ...SwgohGg.getShipList()];
  };

  const getBaseAbilities = (): AbilityDefinition[] => {
    return SwgohGg.getAbilityList();
  };

  const getGuildData = (guild: GuildSettings): GuildData | undefined => {
    // Figure out which data source to use
    if (guild.dataSource === DATASOURCES.SWGOH_GG) {
      // SPREADSHEET.toast(`Get data for guild ${guild.guildId} from swgoh.gg`, 'Refresh data', 5);
      const data = SwgohGg.getGuildData(guild.guildId!);
      if (data) {
        // fix member level
        for (const member of data.members) {
          let level = 1;
          // tslint:disable-next-line: forin
          for (const baseId in member.units) {
            level = Math.max(member.units[baseId].level, level);
          }
          member.level = level;
        }
        return data;
      }
      throw new Error(`Failled: SwgohGg.getGuildData(${guild.guildId})`);
    }
    if (guild.dataSource === DATASOURCES.SWGOH_HELP) {
      // SPREADSHEET.toast(`Get data for guild of ${guild.allyCode} from swgoh.help`, 'Refresh data', 5);
      return SwgohHelp.getGuildData(guild.allyCode!);
    }
  };

  /** check if guilds setup is fresh */
  const isGuildsFresh = (): boolean => {
    return getGuildsHash() === Cache.getGuildsHash();
  };

  /** updates cached guilds setup  */
  const refreshGuilds = (): void => {
    Cache.setGuildsHash(getGuildsHash());
  };

  /** compute hash for SwgohHelp guilds settings */
  const getGuildsHash = (): string => {
    const guilds = SPREADSHEET.getRangeByName(CORERANGE.GUILDDATASOURCESETTINGS).getValues();

    return String(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(guilds)));
  };

  /** check if cached credential hash is fresh */
  const isCredentialFresh = (): boolean => {
    return getCredentialHash() === Cache.getCredentialHash();
  };

  /** updates cached credential hash */
  const refreshCredential = (): void => {
    Cache.setCredentialHash(getCredentialHash());
  };

  /** compute hash for SwgohHelp credential settings */
  const getCredentialHash = (): string => {
    const credential = [
      SPREADSHEET.getRangeByName(CORERANGE.SWGOHHELPUSERNAME).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.SWGOHHELPPASSWORD).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.ADDMEMBERGPTOROSTER).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.ADDMEMBERBATTLESTOROSTER).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.ADDMEMBERGUILDACTIVITIESTOROSTER).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.WHICHUNITTYPE).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.WHICHHEROABILITIES).getValue(),
      SPREADSHEET.getRangeByName(CORERANGE.GETSHIPABILITIES).getValue(),
    ];

    return String(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(credential)));
  };

  /** check if cached rename/add/remove hash is fresh */
  const isRARFresh = (): boolean => {
    return getRARHash() === Cache.getRARHash();
  };

  /** updates cached rename/add/remove hash */
  const refreshRAR = (): void => {
    Cache.setRARHash(getRARHash());
  };

  type rarColumns = [string, string, number, number];

  /** compute hash for Rename/Add/Remove settings */
  const getRARHash = (): string => {
    const rar = (SPREADSHEET.getRangeByName(CORERANGE.RENAMEADDREMOVE).getValues() as rarColumns[])
      .reduce((acc: rarColumns[], e) => {
        if (e[2] > 0 || e[3] > 0) {
          acc.push(e);
        }
        return acc;
      }, [])
      .sort((a, b) => (a[2] !== b[2] ? a[2] - b[2] : a[3] - b[3]));

    return String(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(rar)));
  };

  interface RenameAddRemove {
    rename: Array<{ name: string; allyCode: number }>;
    add: Array<{ guild: string; allyCode: number }>;
    remove: number[];
  }

  /** compute hash for Rename/Add/Remove settings */
  const getRAR = (datas: GuildData[]): RenameAddRemove => {
    const rar = SPREADSHEET.getRangeByName(CORERANGE.RENAMEADDREMOVE)
      .getValues()
      .reduce(
        (acc: RenameAddRemove, e) => {
          const guild = `${e[0]}`.trim();
          const name = `${e[1]}`.trim();
          const add = +e[2];
          const remove = +e[3];
          if (add > 0) {
            if (name.length > 0) {
              acc.rename.push({ name, allyCode: add });
            }
            if (guild.length > 0) {
              acc.add.push({ guild, allyCode: add });
            } else {
              if (
                datas.findIndex(
                  (guildData) => guildData.members.findIndex((member) => member.allyCode === add) > -1,
                ) === -1
              ) {
                acc.add.push({ guild: 'PLAYER', allyCode: add });
              }
            }
          }
          if (remove > 0) {
            acc.remove.push(remove);
          }
          return acc;
        },
        { rename: [], add: [], remove: [] },
      ) as RenameAddRemove;

    rar.remove = rar.remove.unique();

    return rar;
  };

  /** all things cached */
  namespace Cache {
    const cacheKey = SPREADSHEET.getId();
    const seconds = 3600; // 1 hour
    const service = CacheService.getScriptCache();

    // tslint:disable-next-line: no-shadowed-variable
    export const getGuildsHash = (): string => service.get(`${cacheKey}-guilds`);

    export const setGuildsHash = (hash: string): void => service.put(`${cacheKey}-guilds`, hash, seconds);

    // tslint:disable-next-line: no-shadowed-variable
    export const getCredentialHash = (): string => service.get(`${cacheKey}-cred`);

    export const setCredentialHash = (hash: string): void => service.put(`${cacheKey}-cred`, hash, seconds);

    // tslint:disable-next-line: no-shadowed-variable
    export const getRARHash = (): string => service.get(`${cacheKey}-rar`);

    export const setRARHash = (hash: string): void => service.put(`${cacheKey}-rar`, hash, seconds);
  }

  namespace Sheets {
    export const setValues = (sheet: Spreadsheet.Sheet, values: any[][], headers: string[]) => {
      let last: number;
      let max: number;
      values.unshift(headers);

      sheet
        .clear()
        .getRange(1, 1, values.length, headers.length)
        .setValues(values);

      [last, max] = [sheet.getLastRow(), sheet.getMaxRows()];
      if (max > last) {
        sheet.deleteRows(last + 1, max - last);
      }

      [last, max] = [sheet.getLastColumn(), sheet.getMaxColumns()];
      if (max > last) {
        sheet.deleteColumns(last + 1, max - last);
      }
    };

    export const numberOfCells = () => {
      const formatThousandsNoRounding = (n: number, dp?: number) => {
        const s = `${n}`;
        const b = n < 0 ? 1 : 0;
        const i = s.lastIndexOf('.');
        let j = i === -1 ? s.length : i;
        let r = '';
        const d = s.substr(j + 1, dp);
        // tslint:disable-next-line: no-conditional-assignment
        while ((j -= 3) > b) {
          r = `,${s.substr(j, 3)}${r}`;
        }

        return `${s.substr(0, j + 3)}${r}${dp ? `.${d}${d.length < dp ? '00000'.substr(0, dp - d.length) : ''}` : ''}`;
      };

      const sheets = SPREADSHEET.getSheets();
      let cellsCount = 0;
      for (const sheet of sheets) {
        cellsCount += sheet.getMaxColumns() * sheet.getMaxRows();
      }
      // const ratio = cellsCount / 2000000 * 100;
      Logger.log(formatThousandsNoRounding(cellsCount));
    };

    export function getSheetByNameOrDie(name: string) {
      const sheet = SPREADSHEET.getSheetByName(name);
      if (sheet) {
        return sheet;
      } else {
        throw new Error(`Missing sheet labeled '${name}'`);
      }
    }
  }
}

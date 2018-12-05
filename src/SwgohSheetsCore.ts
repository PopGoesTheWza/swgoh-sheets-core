/**
 * @OnlyCurrentDoc
 */

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

type URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

type KeyedType<T> = {
  [key: string]: T;
};

type KeyedNumbers = KeyedType<number>;

interface GuildData {
  id: number;
  name: string;
  members: PlayerData[];
}

interface PlayerData {
  allyCode: number;
  level?: number;
  link?: string;
  name: string;
  gp: number;
  heroesGp: number;
  shipsGp: number;
  fleetArenaBattlesWon: number;
  squadArenaBattlesWon: number;
  normalBattlesWon: number;
  hardBattlesWon: number;
  galacticWarBattlesWon: number;
  guildRaidsWon: number;
  guildTokensEarned: number;
  gearDonatedInGuildExchange: number;
  units: UnitInstances;
}

type UnitsDefinitions = {
  heroes: UnitDefinition[];
  ships: UnitDefinition[];
};

/** A unit's name, baseId and tags */
interface UnitDefinition {
  /** Unit Id */
  baseId: string;
  name: string;
  /** Alignment, role and tags */
  tags: string;
  type?: number;
}

type Ability = {
  name: string;
  type: string;
  tier: number;
  isZeta: boolean;
};

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

type UnitInstances = {
  [key: string]: UnitInstance;
};

/** Constants for data source */
enum DATASOURCES {
  /** Use swgoh.gg API as data source */
  SWGOH_GG = 'SwgohGg',
  /** Use swgoh.help API as data source */
  SWGOH_HELP = 'SwgohHelp',
}

/** Constants for sheets name */
enum SHEETS {
  SETUP = 'coreSetup',
  MEMBERS = 'coreRoster',
  HEROES = 'coreHeroes',
  SHIPS = 'coreShips',
  // UNITS = 'coreUnits',
}

/** settings related functions */
namespace config {

  /** use members galactic power stats */
  export function membersGP(): boolean {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(19, 9)
      .getValue() as string;

    return result === 'ON';
  }

  /** use members galactic power stats */
  export function membersBattles(): boolean {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(20, 9)
      .getValue() as string;

    return result === 'ON';
  }

  /** use members galactic power stats */
  export function membersGuildActivities(): boolean {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(21, 9)
      .getValue() as string;

    return result === 'ON';
  }

  export enum UNIT_TYPES {
    BOTH = 'Heroes & Ships',
    HERO = 'Heroes only',
    SHIP = 'Ships only',
  }

  /** use members galactic power stats */
  export function unitTypes(): string {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(24, 9)
      .getValue() as string;

    return result;
  }

  export enum HERO_ABILITIES {
    NONE = 'None',
    ZETAS = 'Zetas only',
    WORTHY = 'Leader & zetas',
    DETAILED = 'Detailed',
  }

  /** use members galactic power stats */
  export function heroAbilities(): string {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(25, 9)
      .getValue() as string;

    return result;
  }

  /** use members galactic power stats */
  export function shipAbilities(): boolean {

    const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(21, 9)
      .getValue() as string;

    return result === 'ON';
  }

  /** SwgohHelp related settings */
  export namespace SwgohHelp {

    /** Get the SwgohHelp API username */
    export function username(): string {

      const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
        .getRange(15, 9)
        .getValue() as string;

      return result;
    }

    /** Get the SwgohHelp API password */
    export function password(): string {

      const result = SPREADSHEET.getSheetByName(SHEETS.SETUP)
        .getRange(16, 9)
        .getValue() as string;

      return result;
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

  export const refreshData = (): void => {

    const activeGuilds = getActiveGuilds();
    const credentialIsFresh = isCredentialFresh();
    const rarIsFresh = isRARFresh();
    const guildsIsFresh = isGuildsFresh();

    let staleGuilds = guildsIsFresh && credentialIsFresh && rarIsFresh ? [] : activeGuilds;

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

    if (staleGuilds.length > 0) {

      staleGuilds = activeGuilds;  // TODO proper handling of stale

      const baseUnits = getBaseUnits();
      const datas = staleGuilds.map(e => getGuildData(e));
      renameAddRemove(staleGuilds, datas);
      setGuildNames(staleGuilds, datas);
      writeGuildNames(activeGuilds);
      writeRoster(datas);
      writeUnits(datas, baseUnits);
      refreshGuilds();
      refreshCredential();
      refreshRAR();
    } else {
      const UI = SpreadsheetApp.getUi();
      UI.alert(
        'Nothing to refresh',
        'Data from API is cached for one hour',
        UI.ButtonSet.OK,
      );

      return;
    }
  };

  const renameAddRemove = (settings: GuildSettings[], datas: GuildData[]): void => {
    const rar = getRAR();
    for (const add of rar.add) {
      const guild = add.guild;
      const target = datas.find(e => e.name === guild);
      if (target) {
        const allyCode = add.allyCode;
        // is in another guild?
        let found = false;
        for (const source of datas) {
          const i = source.members.findIndex(e => e.allyCode === allyCode);
          found = i > -1;
          if (found) {
            if (source !== target) {
              target.members.push(source.members.splice(i, 1)[0]);
            }
            break;
          }
        }
        if (!found) {
          // TODO: complete Add feature
          // try to guess data source
          // const setup = settings.find(e => e.guildId === target.id);
          // const g = SwgohGg.getPlayerData(allyCode);
          // const h = SwgohHelp.getPlayerData(allyCode);
        }
      }
    }
    for (const rename of rar.rename) {
      const allyCode = rename.allyCode;
      for (const data of datas) {
        const i = data.members.findIndex(e => e.allyCode === allyCode);
        if (i > -1) {
          data.members[i].name = rename.name;
        }
      }
    }
    for (const allyCode of rar.remove) {
      for (const data of datas) {
        const i = data.members.findIndex(e => e.allyCode === allyCode);
        if (i > -1) {
          data.members.splice(i, 1);
        }
      }
    }
  };

  const setGuildNames = (settings: GuildSettings[], datas: GuildData[]): void => {
    for (const i in settings) {
      const setup = settings[i];
      if (setup.name.length === 0) {
        setup.name = datas[i].name;
      }
    }
    // TODO: duplicates guild name
  };

  const writeGuildNames = (settings: GuildSettings[]): void => {

    const range = SPREADSHEET.getSheetByName(SHEETS.SETUP).getRange(3, 2, 10, 5);
    const values = range.getValues() as string[][];
    for (const setting of settings) {
      const name = setting.name;
      const dataSource = setting.dataSource;
      let row;
      if (dataSource === DATASOURCES.SWGOH_GG) {
        const id = setting.guildId;
        row = values.find(e => e[0] === 'ON' && e[2] === dataSource && +e[3] === id);
      } else {
        const id = setting.allyCode;
        row = values.find(e => e[0] === 'ON' && e[2] === dataSource && +e[4] === id);
      }
      if (row && row[1] !== name) {
        row[1] = name;
      }
    }
    range.setValues(values);
  };

  const writeRoster = (datas: GuildData[]): void => {
    const addGP = config.membersGP();
    const addBattles = config.membersBattles();
    const addActivities = config.membersGuildActivities();
    const values: any[][] = [];
    for (const data of datas) {
      if (data) {
        const guildId = data.id;
        const guildName = data.name;
        for (const member of data.members) {
          let columns = [
            guildId,
            guildName,
            member.name,
            member.allyCode,
            member.level,
          ];
          if (addGP) {
            columns = columns.concat([
              member.gp,
              member.heroesGp,
              member.shipsGp,
            ]);
          }
          if (addBattles) {
            columns = columns.concat([
              member.fleetArenaBattlesWon,
              member.squadArenaBattlesWon,
              member.normalBattlesWon,
              member.hardBattlesWon,
              member.galacticWarBattlesWon,
            ]);
          }
          if (addActivities) {
            columns = columns.concat([
              member.guildRaidsWon,
              member.guildTokensEarned,
              member.gearDonatedInGuildExchange,
            ]);
          }
          values.push(columns);
        }
      }
    }
    let headers = [
      'guildId',
      'guildName',
      'name',
      'allyCode',
      'level',
    ];

    if (addGP) {
      headers = headers.concat([
        'gp',
        'heroesGp',
        'shipsGp',
      ]);
    }
    if (addBattles) {
      headers = headers.concat([
        'fleetArenaBattlesWon',
        'squadArenaBattlesWon',
        'normalBattlesWon',
        'hardBattlesWon',
        'galacticWarBattlesWon',
      ]);
    }
    if (addActivities) {
      headers = headers.concat([
        'guildRaidsWon',
        'guildTokensEarned',
        'gearDonatedInGuildExchange',
      ]);
    }
    Sheets.setValues(
      SPREADSHEET.getSheetByName(SHEETS.MEMBERS),
      values,
      headers,
    );
  };

  const writeUnits = (datas: GuildData[], baseUnits: UnitDefinition[]): void => {
    // TODO: use distinct sheet for base unit data (tags, abilities...)
    const unitTypes = config.unitTypes();
    const addHeroes = unitTypes !== config.UNIT_TYPES.SHIP;
    const addShips = unitTypes !== config.UNIT_TYPES.HERO;
    const heroAbilities = config.heroAbilities();
    const shipAbilities = config.shipAbilities();
    const heroes: any[][] = [];
    const ships: any[][] = [];
    // const values: any[][] = [];
    const heroesAbilities = { count: 0, columns: 0 };
    const shipsAbilities = { count: 0, columns: 0 };
    // let nAbilities = 0;
    // let nColumns = 0;
    for (const data of datas) {
      if (data) {
        // const guildId = data.id;
        // const guildName = data.name;
        for (const member of data.members) {
          const units = member.units;
          for (const baseId in units) {
            const def = baseUnits.find(e => e.baseId === baseId);
            const unit = units[baseId];
            const type = unit.type;
            if ((addHeroes && type === Units.TYPES.HERO)
              || (addShips && type === Units.TYPES.SHIP)) {
              const columns = type === Units.TYPES.HERO ? [
                member.name,
                member.allyCode,
                def.name,
                def.tags,
                unit.rarity,
                unit.level,
                unit.gearLevel,
                unit.power,
              ] : [
                member.name,
                member.allyCode,
                def.name,
                def.tags,
                unit.rarity,
                unit.level,
                unit.power,
              ];
              if (heroAbilities !== config.HERO_ABILITIES.NONE && type === Units.TYPES.HERO) {
                // TODO: write abilites (enum UnitsAbilities)
                if (heroAbilities === config.HERO_ABILITIES.ZETAS) {
                  let count = 0;
                  for (const ability of unit.abilities) {
                    if (ability.isZeta && ability.tier === 8) {
                      columns.push(ability.name);
                      columns.push(ability.type);
                      count += 1;
                    }
                  }
                  heroesAbilities.count = Math.max(count, heroesAbilities.count);
                } else if (heroAbilities === config.HERO_ABILITIES.WORTHY) {
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
                } else if (heroAbilities === config.HERO_ABILITIES.DETAILED) {
                  for (const ability of unit.abilities) {
                    columns.push(ability.name);
                    columns.push(ability.type);
                    columns.push(ability.isZeta ? 'true' : 'false');
                    columns.push(ability.tier);
                  }
                  heroesAbilities.count = Math.max(unit.abilities.length, heroesAbilities.count);
                }
                heroesAbilities.columns = Math.max(columns.length, heroesAbilities.columns);
                // nAbilities = Math.max(unit.abilities.length, nAbilities);
                // nColumns = Math.max(columns.length, nColumns);
              } else if (shipAbilities && type === Units.TYPES.SHIP) {
                // TODO: write abilites (enum UnitsAbilities)
                for (const ability of unit.abilities) {
                  columns.push(ability.name);
                  columns.push(ability.type);
                  columns.push(ability.tier);
                }
                shipsAbilities.count = Math.max(unit.abilities.length, shipsAbilities.count);
                shipsAbilities.columns = Math.max(columns.length, shipsAbilities.columns);
                // nAbilities = Math.max(unit.abilities.length, nAbilities);
                // nColumns = Math.max(columns.length, nColumns);
              }
              (type === Units.TYPES.HERO ? heroes : ships).push(columns);
              // values.push(columns);
            }
          }
        }
      }
    }
    const heroesHeaders = [
      'name',
      'allyCode',
      'unit',
      'tags',
      'rarity',
      'level',
      'gearLevel',
      'power',
    ];
    const shipsHeaders = [
      'name',
      'allyCode',
      'unit',
      'tags',
      'rarity',
      'level',
      'power',
    ];
    // TODO: write abilites (enum UnitsAbilities)
    if (heroAbilities === config.HERO_ABILITIES.ZETAS) {
      for (let i = 0; i < heroesAbilities.count; i += 1) {
        heroesHeaders.push(`ability_${i}`);
        heroesHeaders.push(`type_${i}`);
      }
    } else if (heroAbilities === config.HERO_ABILITIES.WORTHY
      || heroAbilities === config.HERO_ABILITIES.DETAILED) {
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
    // for (let i = 0; i < nAbilities; i += 1) {
    //   headers.push(`ability_${i}`);
    //   headers.push(`type_${i}`);
    //   headers.push(`isZeta_${i}`);
    //   headers.push(`tier_${i}`);
    // }
    Sheets.setValues(
      SPREADSHEET.getSheetByName(SHEETS.HEROES),
      heroes.map(e =>
        e.length !== heroesAbilities.columns
          ? e.concat(Array(heroesAbilities.columns).fill(null)).slice(0, heroesAbilities.columns)
          : e),
      heroesHeaders,
    );
    Sheets.setValues(
      SPREADSHEET.getSheetByName(SHEETS.SHIPS),
      ships.map(e =>
        e.length !== shipsAbilities.columns
          ? e.concat(Array(shipsAbilities.columns).fill(null)).slice(0, shipsAbilities.columns)
          : e),
      shipsHeaders,
    );
    // Sheets.setValues(
    //   SPREADSHEET.getSheetByName(SHEETS.UNITS),
    //   values.map(e =>
    //     e.length !== nColumns ? e.concat(Array(nColumns).fill(null)).slice(0, nColumns) : e),
    //   headers,
    // );
  };

  type GuildSettings = {
    name: string,
    dataSource: DATASOURCES,
    guildId?: number,
    allyCode?: number,
  };

  /** returns an array of active guild settings */
  const getActiveGuilds = (): GuildSettings[] => {

    const guilds = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(3, 2, 10, 5).getValues().reduce(
      (acc: GuildSettings[], columns: string[]) => {

        if (columns[0] === 'ON') {  // hardcoded checkbox value
          const name = columns[1];
          const dataSource = columns[2] as DATASOURCES;
          if (dataSource === DATASOURCES.SWGOH_GG) {
            acc.push({ name, dataSource, guildId: +columns[3] });
          } else if (dataSource === DATASOURCES.SWGOH_HELP) {
            acc.push({ name, dataSource, allyCode: +columns[4] });
          }
        }

        return acc;
      },
      [],
    ) as GuildSettings[];

    return guilds;
  };

  const getBaseUnits = (): UnitDefinition[] => {
    return SwgohGg.getHeroList().concat(SwgohGg.getShipList());
  };

  const getGuildData = (guild: GuildSettings): GuildData => {

    // Figure out which data source to use
    if (guild.dataSource === DATASOURCES.SWGOH_GG) {
      const data = SwgohGg.getGuildData(guild.guildId);
      // fix member level
      for (const member of data.members) {
        let level = 1;
        for (const baseId in member.units) {
          level = Math.max(member.units[baseId].level, level);
        }
        member.level = level;
      }
      return data;
    }
    if (guild.dataSource === DATASOURCES.SWGOH_HELP) {
      return SwgohHelp.getGuildData(guild.allyCode);
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

    const guilds = SPREADSHEET.getSheetByName(SHEETS.SETUP)
      .getRange(3, 2, 10, 5).getValues();

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(guilds),
    ));
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

    const sheet = SPREADSHEET.getSheetByName(SHEETS.SETUP);

    const credential = sheet.getRange(15, 9, 2).getValues().concat(
      sheet.getRange(19, 9, 3).getValues(),
      sheet.getRange(24, 9, 3).getValues());

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(credential),
    ));
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

    const sheet = SPREADSHEET.getSheetByName(SHEETS.SETUP);

    const rar = sheet
      .getRange(15, 3, sheet.getMaxRows(), 4).getValues()
      .reduce(
        (acc: rarColumns[], e: rarColumns) => {
          if (e[2] > 0 || e[3] > 0) acc.push(e);
          return acc;
        },
        [],
      )
      .sort((a, b) => a[2] !== b[2] ? a[2] - b[2] : a[3] - b[3]) as rarColumns[];

    return String(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(rar),
    ));
  };

  type RenameAddRemove = {
    rename: { name: string, allyCode: number }[],
    add: { guild: string, allyCode: number }[],
    remove: number[],
  };

  /** compute hash for Rename/Add/Remove settings */
  const getRAR = (): RenameAddRemove => {

    const sheet = SPREADSHEET.getSheetByName(SHEETS.SETUP);

    const rar = sheet
      .getRange(15, 3, sheet.getMaxRows(), 4).getValues()
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
    const seconds = 3600;  // 1 hour
    const service = CacheService.getScriptCache();

    export const getGuildsHash = (): string =>
      service.get(`${cacheKey}-guilds`);

    export const setGuildsHash = (hash: string): void =>
      service.put(`${cacheKey}-guilds`, hash, seconds);

    export const getCredentialHash = (): string =>
      service.get(`${cacheKey}-cred`);

    export const setCredentialHash = (hash: string): void =>
      service.put(`${cacheKey}-cred`, hash, seconds);

    export const getRARHash = (): string =>
      service.get(`${cacheKey}-rar`);

    export const setRARHash = (hash: string): void =>
      service.put(`${cacheKey}-rar`, hash, seconds);

  }

  namespace Sheets {

    export const setValues = (sheet: GoogleAppsScript.Spreadsheet.Sheet, values, headers) => {

      values.unshift(headers);

      sheet.clear().getRange(1, 1, values.length, headers.length)
        .setValues(values);

      const lastRow = sheet.getLastRow();
      const maxRows = sheet.getMaxRows();
      if (maxRows > lastRow) {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      }

      const lastCol = sheet.getLastColumn();
      const maxcols = sheet.getMaxColumns();
      if (maxcols > lastCol) {
        sheet.deleteColumns(lastCol + 1, maxcols - lastCol);
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
        while ((j -= 3) > b) {
          r = `,${s.substr(j, 3)}${r}`;
        }

        // tslint:disable-next-line:max-line-length
        return `${s.substr(0, j + 3)}${r}${dp ? `.${d}${d.length < dp ? ('00000').substr(0, dp - d.length) : ''}` : ''}`;
      };

      const sheets = SPREADSHEET.getSheets();
      let cellsCount = 0;
      for (const sheet of sheets) {
        cellsCount += (sheet.getMaxColumns() * sheet.getMaxRows());
      }
      // const ratio = cellsCount / 2000000 * 100;
      Logger.log(formatThousandsNoRounding(cellsCount));
    };

  }

}

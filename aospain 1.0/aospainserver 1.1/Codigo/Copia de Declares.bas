Attribute VB_Name = "Declaraciones"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public MixedKey As Long
Public ServerIp As String
Public CrcSubKey As String

Type tEstadisticasDiarias
    Segundos As Double
    MaxUsuarios As Integer
    Promedio As Integer
End Type
    
Public DayStats As tEstadisticasDiarias

Public aDos As New clsAntiDoS
Public aClon As New clsAntiMassClon
Public TrashCollector As New Collection


Public Const MAXSPAWNATTEMPS = 60
Public Const MAXUSERMATADOS = 9000000
Public Const LoopAdEternum = 999
Public Const FXSANGRE = 14


Public Const iFragataFantasmal = 87

Public Type tLlamadaGM
    Usuario As String * 255
    Desc As String * 255
End Type

Public Const LimiteNewbie = 8

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Const NingunEscudo = 2
Public Const NingunCasco = 2

Public Const EspadaMataDragonesIndex = 402

Public Const MAXMASCOTASENTRENADOR = 7

Public Const FXWARP = 1
Public Const FXCURAR = 2

Public Const FXMEDITARCHICO = 17
Public Const FXMEDITARMEDIANO = 5
Public Const FXMEDITARMEDIANODOS = 4
Public Const FXMEDITARGRANDE = 6
Public Const FXMEDITARPOWA = 18

Public Const POSINVALIDA = 3

Public Const Bosque = "BOSQUE"
Public Const Nieve = "NIEVE"
Public Const Desierto = "DESIERTO"

Public Const Ciudad = "CIUDAD"
Public Const Campo = "CAMPO"
Public Const Dungeon = "DUNGEON"

' <<<<<< Targets >>>>>>
Public Const uUsuarios = 1
Public Const uNPC = 2
Public Const uUsuariosYnpc = 3
Public Const uTerreno = 4

' <<<<<< Acciona sobre >>>>>>
Public Const uPropiedades = 1
Public Const uEstado = 2
Public Const uInvocacion = 4
Public Const uMaterializa = 3

Public Const DRAGON = 6
Public Const MATADRAGONES = 1

Public Const MAX_MENSAJES_FORO = 35

Public Const MAXUSERHECHIZOS = 35


Public Const EsfuerzoTalarGeneral = 4
Public Const EsfuerzoTalarLe�ador = 2

Public Const EsfuerzoPescarPescador = 1
Public Const EsfuerzoPescarGeneral = 3

Public Const EsfuerzoExcavarMinero = 2
Public Const EsfuerzoExcavarGeneral = 5


Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const Guardias = 6

Public Const MAXREP = 6000000
Public Const MAXORO = 90000000
Public Const MAXEXP = 99999999

Public Const MAXATRIBUTOS = 35
Public Const MINATRIBUTOS = 6

Public Const LingoteHierro = 386
Public Const LingotePlata = 387
Public Const LingoteOro = 388
Public Const Le�a = 58


Public Const MAXNPCS = 10000
Public Const MAXCHARS = 10000

Public Const HACHA_LE�ADOR = 127
Public Const PIQUETE_MINERO = 187

Public Const DAGA = 15
Public Const FOGATA_APAG = 136
Public Const FOGATA = 63
Public Const ORO_MINA = 194
Public Const PLATA_MINA = 193
Public Const HIERRO_MINA = 192
Public Const MARTILLO_HERRERO = 389
Public Const SERRUCHO_CARPINTERO = 198
Public Const ObjArboles = 4

Public Const NPCTYPE_COMUN = 0
Public Const NPCTYPE_REVIVIR = 1
Public Const NPCTYPE_GUARDIAS = 2
Public Const NPCTYPE_ENTRENADOR = 3
Public Const NPCTYPE_BANQUERO = 4



Public Const FX_TELEPORT_INDEX = 1


Public Const MIN_APU�ALAR = 10

'********** CONSTANTANTES ***********
Public Const NUMSKILLS = 21
Public Const NUMATRIBUTOS = 5
Public Const NUMCLASES = 17
Public Const NUMRAZAS = 7

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4


Public Const MAXMASCOTAS = 3

'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const vlASALTO = 100
Public Const vlASESINO = 1000
Public Const vlCAZADOR = 5
Public Const vlNoble = 5
Public Const vlLadron = 25
Public Const vlProleta = 2



'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const iCuerpoMuerto = 8
Public Const iCabezaMuerto = 500


Public Const iORO = 12
Public Const Pescado = 139


'%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
Public Const Suerte = 1
Public Const Magia = 2
Public Const Robar = 3
Public Const Tacticas = 4
Public Const Armas = 5
Public Const Meditar = 6
Public Const Apu�alar = 7
Public Const Ocultarse = 8
Public Const Supervivencia = 9
Public Const Talar = 10
Public Const Comerciar = 11
Public Const Defensa = 12
Public Const Pesca = 13
Public Const Mineria = 14
Public Const Carpinteria = 15
Public Const Herreria = 16
Public Const Liderazgo = 17
Public Const Domar = 18
Public Const Proyectiles = 19
Public Const Wresterling = 20
Public Const Navegacion = 21

Public Const FundirMetal = 88

Public Const XA = 40
Public Const XD = 10
Public Const Balance = 9

Public Const Fuerza = 1
Public Const Agilidad = 2
Public Const Inteligencia = 3
Public Const Carisma = 4
Public Const Constitucion = 5


Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel
Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLe�ador = 23
Public Const AdicionalSTPescador = 20
Public Const AdicionalSTMinero = 25

'Tama�o del mapa
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

'Tama�o del tileset
Public Const TileSizeX = 32
Public Const TileSizeY = 32

'Tama�o en Tiles de la pantalla de visualizacion
Public Const XWindow = 17
Public Const YWindow = 13

'Sonidos
Public Const SOUND_BUMP = 1
Public Const SOUND_SWING = 2
Public Const SOUND_TALAR = 13
Public Const SOUND_PESCAR = 14
Public Const SOUND_MINERO = 15
Public Const SND_WARP = 3
Public Const SND_PUERTA = 5
Public Const SOUND_NIVEL = 6
Public Const SOUND_COMIDA = 7
Public Const SND_USERMUERTE = 11
Public Const SND_IMPACTO = 10
Public Const SND_IMPACTO2 = 12
Public Const SND_LE�ADOR = 13
Public Const SND_FOGATA = 14
Public Const SND_AVE = 21
Public Const SND_AVE2 = 22
Public Const SND_AVE3 = 34
Public Const SND_GRILLO = 28
Public Const SND_GRILLO2 = 29
Public Const SOUND_SACARARMA = 25
Public Const SND_ESCUDO = 37
Public Const MARTILLOHERRERO = 41
Public Const LABUROCARPINTERO = 42
Public Const SND_CREACIONCLAN = 44
Public Const SND_ACEPTADOCLAN = 43
Public Const SND_DECLAREWAR = 45
Public Const SND_BEBER = 46

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20

'<------------------CATEGORIAS PRINCIPALES--------->
Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_CONTENEDORES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_FOROS = 10
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LE�A = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_TELEPORT = 19
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_CUALQUIERA = 1000
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35

'<------------------SUB-CATEGORIAS----------------->
Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_CA�A = 138



'Tipo de posicones
'1 Modifica la Agilidad
'2 Modifica la Fuerza
'3 Repone HP
'4 Repone Mana

'Texto
Public Const FONTTYPE_TALK = "~255~255~255~0~0"
Public Const FONTTYPE_FIGHT = "~255~0~0~1~0"
Public Const FONTTYPE_WARNING = "~32~51~223~1~1"
Public Const FONTTYPE_INFO = "~65~190~156~0~0"
Public Const FONTTYPE_VENENO = "~0~255~0~0~0"
Public Const FONTTYPE_GUILD = "~255~255~255~1~0"

'Estadisticas
Public Const STAT_MAXELV = 99
Public Const STAT_MAXHP = 999
Public Const STAT_MAXSTA = 999
Public Const STAT_MAXMAN = 2000
Public Const STAT_MAXHIT = 99
Public Const STAT_MAXDEF = 99

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Const SND_NODEFAULT = &H2

Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10



'**************************************************************
'**************************************************************
'************************ TIPOS *******************************
'**************************************************************
'**************************************************************

Type tHechizo
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As Byte
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    
    Invisibilidad As Byte
    Paraliza As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    
    Invoca As Byte
    NumNpc As Integer
    Cant As Integer
    
    Materializa As Byte
    ItemIndex As Byte
    
    MinSkill As Integer
    ManaRequerido As Integer

    Target As Byte
End Type

Type LevelSkill

LevelValue As Integer

End Type

Type UserOBJ
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    HerramientaEqpObjIndex As Integer
    HerramientaEqpSlot As Integer
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    NroItems As Integer
End Type


Type Position
    X As Integer
    Y As Integer
End Type

Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Type FXdata
    Nombre As String
    GrhIndex As Integer
    Delay As Integer
End Type

'Datos de user o npc
Type Char
    CharIndex As Integer
    Head As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    loops As Integer
    
    Heading As Byte
End Type

'Tipos de objetos
Public Type ObjData
    
    Name As String 'Nombre del obj
    
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    SubTipo As Integer 'Tipo enum que determina cuales son las caract del obj
    
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    
    Respawn As Byte
    
    'Solo contenedores
    MaxItems As Integer
    Conte As Inventario
    Apu�ala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer ' Minimo puntos de vida
    MaxHP As Integer ' Maximo puntos de vida
    
    
    MineralIndex As Integer
    LingoteInex As Integer
    
    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    'Puntos de Stamina que da
    MinSta As Integer ' Minimo puntos de stamina
    
    'Pociones
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHIT As Integer 'Minimo golpe
    MaxHIT As Integer 'Maximo golpe
    
    MinHam As Integer
    MinSed As Integer
    
    Def As Integer
    MinDef As Integer ' Armaduras
    MaxDef As Integer ' Armaduras
    
    Ropaje As Integer 'Indice del grafico del ropaje
    
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    CascoAnim As Integer
    
    Valor As Long     ' Precio
    
    Cerrada As Integer
    Llave As Byte
    Clave As Long 'si clave=llave la puerta se abre o cierra
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    RazaHobbit As Byte
    Mujer As Byte
    Hombre As Byte
    Envenena As Byte
    
    Resistencia As Long
    Agarrable As Byte
    
    
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    
    SkHerreria As Integer
    SkCarpinteria As Integer
    
    texto As String
    
    'Clases que no tienen permitido usar este obj
    ClaseProhibida(1 To NUMCLASES) As String
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    MinInt As Integer
    
    Real As Integer
    Caos As Integer
    
End Type

Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

'[KEVIN]
'Banco Objs
Public Const MAX_BANCOINVENTORY_SLOTS = 40
'[/KEVIN]

'[KEVIN]
Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    NroItems As Integer
End Type
'[/KEVIN]


'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'******* T I P O S   D E    U S U A R I O S **************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Type tReputacion 'Fama del usuario
    NobleRep As Double
    BurguesRep As Double
    PlebeRep As Double
    LadronesRep As Double
    BandidoRep As Double
    AsesinoRep As Double
    Promedio As Double
End Type



'Estadisticas de los usuarios
Type UserStats
    GLD As Long 'Dinero
    Banco As Long
    MET As Integer
    
    MaxHP As Integer
    MinHP As Integer
    
    FIT As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    
    MaxHam As Integer
    MinHam As Integer
    
    MaxAGU As Integer
    MinAGU As Integer
        
    Def As Integer
    Exp As Double
    ELV As Long
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Integer
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Integer
    CriminalesMatados As Integer
    NPCsMuertos As Integer
    
    SkillPts As Integer
    
End Type

'Flags
Type UserFlags
    Muerto As Byte '�Esta muerto?
    Escondido As Byte '�Esta escondido?
    Comerciando As Boolean '�Esta comerciando?
    UserLogged As Boolean '�Esta online?
    Meditando As Boolean
    ModoCombate As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeAtacar As Byte
    PuedeMoverse As Byte
    PuedeLanzarSpell As Byte
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    
    Vuela As Byte
    Navegando As Byte
    Seguro As Boolean
    
    DuracionEfecto As Long
    TargetNpc As Integer ' Npc se�alado por el usuario
    TargetNpcTipo As Integer ' Tipo del npc se�alado
    NpcInv As Integer
    
    Ban As Byte
    AdministrativeBan As Byte
    
    TargetUser As Integer ' Usuario se�alado
    
    TargetObj As Integer ' Obj se�alado
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    
    StatsChanged As Byte
    Privilegios As Byte
    
    ValCoDe As Integer
    
    LastCrimMatado As String
    LastCiudMatado As String
    
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    
End Type

Type UserCounters
    IdleCount As Long
    AttackCounter As Integer
    HPCounter As Integer
    STACounter As Integer
    Frio As Integer
    COMCounter As Integer
    AGUACounter As Integer
    Veneno As Integer
    Paralisis As Integer
    Ceguera As Integer
    Estupidez As Integer
    Invisibilidad As Integer
    PiqueteC As Long
    Pena As Long
    SendMapCounter As WorldPos
    Pasos As Integer
End Type

Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Double
    CiudadanosMatados As Double
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
End Type

Type tGuild
    GuildName As String
    Solicitudes As Long
    SolicitudesRechazadas As Long
    Echadas As Long
    VecesFueGuildLeader As Long
    YaVoto As Byte
    EsGuildLeader As Byte
    FundoClan As Byte
    ClanFundado As String
    ClanesParticipo As Long
    GuildPoints As Double
End Type

'Tipo de los Usuarios
Type User
    
    Name As String
    ID As Long
    
    modName As String
    Password As String
    
    Char As Char 'Define la apariencia
    OrigChar As Char
    
    Desc As String ' Descripcion
    Clase As String
    Raza As String
    Genero As String
    Email As String
    Hogar As String
    
    
    Invent As Inventario
    
    Pos As WorldPos
    
    
    ConnID As Integer 'ID
    RDBuffer As String 'Buffer roto
    
    CommandsBuffer As New CColaArray
    
    '[KEVIN]
    BancoInvent As BancoInventario
    '[/KEVIN]
    
    
    Counters As UserCounters
    
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMacotas As Integer
    
    Stats As UserStats
    Flags As UserFlags
    NumeroPaquetesPorMiliSec As Long
    BytesTransmitidosUser As Long
    BytesTransmitidosSvr As Long
    
    Reputacion As tReputacion
    
    Faccion As tFacciones
    GuildInfo As tGuild
    GuildRef  As cGuild
    
    PrevCRC As Long
    PacketNumber As Long
    RandKey As Long
    
    ip As String
    
     '[Alejo]
    ComUsu As tCOmercioUsuario
    '[/Alejo]
End Type




'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************
'**  T I P O S   D E    N P C S **************************
'*********************************************************
'*********************************************************
'*********************************************************
'*********************************************************

Type NPCStats
    Alineacion As Integer
    MaxHP As Long
    MinHP As Long
    MaxHIT As Integer
    MinHIT As Integer
    Def As Integer
    UsuariosMatados As Integer
    ImpactRate As Integer
End Type

Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Long
    
End Type

Type NPCFlags
    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean '�Esta vivo?
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    
    OldMovement As Byte
    OldHostil As Byte
    
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    Attacking As Integer
    AttackedBy As String
    Category1 As String
    Category2 As String
    Category3 As String
    Category4 As String
    Category5 As String
    BackUp As Byte
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Snd4 As Integer
    
End Type

Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

'<--------- New type for holding the pathfinding info ------>
Type NpcPathFindingInfo
    Path() As tVertice      ' This array holds the path
    Target As Position      ' The location where the NPC has to go
    PathLenght As Integer   ' Number of steps *
    CurPos As Integer       ' Current location of the npc
    TargetUser As Integer   ' UserIndex chased
    NoPath As Boolean       ' If it is true there is no path to the target location
    
    '* By setting PathLenght to 0 we force the recalculation
    '  of the path, this is very useful. For example,
    '  if a NPC or a User moves over the npc's path, blocking
    '  its way, the function NpcLegalPos set PathLenght to 0
    '  forcing the seek of a new path.
    
End Type
'<--------- New type for holding the pathfinding info ------>


Type Npc
    Name As String
    Char As Char 'Define como se vera
    Desc As String
    
    NPCtype As Integer
    Numero As Integer
    
    Level As Integer
    
    InvReSpawn As Byte
    
    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer
    
    Veneno As Byte
    
    Pos As WorldPos 'Posicion
    Orig As WorldPos
    SkillDomar As Integer
    
    Movement As Integer
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    
    Inflacion As Long
    
    GiveEXP As Long
    GiveGLD As Long
    
    Stats As NPCStats
    Flags As NPCFlags
    Contadores As NpcCounters
    
    Invent As Inventario
    CanAttack As Byte
    
    NroExpresiones As Byte
    Expresiones() As String ' le da vida ;)
    
    NroSpells As Byte
    Spells() As Integer  ' le da vida ;)
    
    '<<<<Entrenadores>>>>>
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    
    '<---------New!! Needed for pathfindig----------->
    PFINFO As NpcPathFindingInfo

    
End Type

'**********************************************************
'**********************************************************
'******************** Tipos del mapa **********************
'**********************************************************
'**********************************************************
'Tile
Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Integer
    UserIndex As Integer
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Trigger As Integer
End Type

'Info del mapa
Type MapInfo
    NumUsers As Integer
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Pk As Boolean
    
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
End Type



'********** V A R I A B L E S     P U B L I C A S ***********

Public SERVERONLINE As Boolean
Public ULTIMAVERSION As String
Public BackUp As Boolean

Public ListaRazas() As String
Public SkillsNames() As String
Public ListaClases() As String


Public ENDL As String
Public ENDC As String

Public recordusuarios As Long

'Directorios
Public IniPath As String
Public CharPath As String
Public MapPath As String
Public DatPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public ResPos As WorldPos
Public StartPos As WorldPos 'Posicion de comienzo


Public NumUsers As Integer 'Numero de usuarios actual
Public LastUser As Integer
Public LastChar As Integer
Public NumChars As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumFX As Integer
Public NumMaps As Integer
Public NumObjDatas As Integer
Public NumeroHechizos As Integer
Public AllowMultiLogins As Byte
Public IdleLimit As Integer
Public MaxUsers As Integer
Public HideMe As Byte
Public LastBackup As String
Public Minutos As String
Public haciendoBK As Boolean
Public Oscuridad As Integer
Public NocheDia As Integer


'*****************ARRAYS PUBLICOS*************************
Public UserList() As User 'USUARIOS
Public Npclist() As Npc 'NPCS
Public MapData() As MapBlock
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public CharList() As Integer
Public ObjData() As ObjData
Public FX() As FXdata
Public SpawnList() As tCriaturasEntrenador
Public LevelSkill(1 To 50) As LevelSkill
Public ForbidenNames() As String
Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ObjCarpintero() As Integer

'*********************************************************

Public Nix As WorldPos
Public Ullathorpe As WorldPos
Public Banderbill As WorldPos
Public Lindos As WorldPos

Public Prision As WorldPos
Public Libertad As WorldPos


Public Ayuda As New cCola


Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GenCrC Lib "crc" Alias "GenCrc" (ByVal CrcKey As Long, ByVal CrcString As String) As Long


Sub PlayWaveAPI(file As String)

On Error Resume Next
Dim rc As Integer

rc = sndPlaySound(file, SND_ASYNC)

End Sub


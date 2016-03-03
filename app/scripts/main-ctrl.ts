namespace app {
  'use strict'

  class Sheet {
    constructor(
      public name: string
    ) {}
    public data: string[][] = []
    public active: boolean = false
  }

  interface IMainScope extends angular.IScope {
    loadFile: (file: File) => void
    loadSheet: () => void
    saveCSVFile: () => void
    fileName: string
    sheets: Sheet[]
    originalHeadings: string[]
    originalData: string[][]
    error: IError
    skipRows: number
    output: Table
    firstRowIsHeader: boolean
    selectedHeading: number
    deleteOutputColumn: (column: string) => void
    deleteOutputMapping: (rowIndex: number) => void
    addHeading: () => void
    addMapping: (key: string, column: number, origColumn: string) => void
    addMappingKey: (key: string) => void
    deleteMapping: (key: string, column: number, origColumnIndex: number) => void
    updateOutput: () => void
    selectMappingToLoad: () => void
    loadMapping: (mapping: Table) => void
    showSaveMappingDialog: () => void
    saveMapping: (name: string) => void
    deleteSavedMapping: (name: string) => void
    importMapping: (file: File) => void
    exportMapping: () => void
    mappings: {[id: string]: Table}
  }

  class Table {
    public headings: Heading[] = [new Heading()]
    public mappings: {[id: number]: number[][]} = {}
    public data: string[][]
  }

  class Heading {
    public type: string = 'Distinct values'
    public label: string
  }

  interface IError {
    data?: string
    config?: {}
    status?: number
  }

  declare var XLSX: any

  export class MainController {
    constructor(private $scope: IMainScope, $uibModal: angular.ui.bootstrap.IModalService, $localStorage: any) {
      if (!$localStorage.mappings) $localStorage.mappings = {}
      $scope.mappings = $localStorage.mappings
      let updateOutput: () => void = $scope.updateOutput = () => {
        $scope.output.data = []
        let idRowMapMap: {[id: number]: {[id: string]: number[]}} = {}
        let idSet: {[id: string]: boolean} = {}
        for (let idS in $scope.output.mappings) {
          /* tslint:disable:no-var-keyword */
          var id: number = parseInt(idS, 10)
          var idRowMap: {[id: string]: number[]} = {}
          /* tslint:enable:no-var-keyword */
          idRowMapMap[id] = idRowMap
          $scope.originalData.forEach((row: string[], index: number) => {
            if (!idRowMap[row[id]]) idRowMap[row[id]] = []
            idRowMap[row[id]].push(index)
            idSet[row[id]] = true
          })
          for (let nid in idSet) $scope.output.data.push([nid])
        }
        /* tslint:disable:no-var-keyword */
        for (var i: number = 1; i < $scope.output.headings.length; i++) {
          /* tslint:enable:no-var-keyword */
          $scope.output.data.forEach((row: string[]) => {
            /* tslint:disable:no-var-keyword */
            var nid: string = row[0]
            let values: {[id: string]: boolean} = {}
            for (var idS in $scope.output.mappings) if ($scope.output.mappings[idS][i] && idRowMapMap[idS][nid]) {
              /* tslint:enable:no-var-keyword */
              $scope.output.mappings[idS][i].forEach(originalColumn => idRowMapMap[idS][nid].forEach(originalRow => { if ($scope.originalData[originalRow][originalColumn]) values[$scope.originalData[originalRow][originalColumn]] = true }))
            }
            let valueString: string = undefined
            let heading: Heading = $scope.output.headings[i]
            switch (heading.type) {
              case 'Distinct values':
                valueString = ''
                for (let val in values) valueString += val + '; '
                valueString = valueString.substring(0, valueString.length - 2)
                if (valueString.length === 0) valueString = null
                break
              case 'Distinct count':
                let count: number = 0
                /* tslint:disable:no-unused-variable */
                for (let val in values) count++
                /* tslint:enable:no-unused-variable */
                valueString = '' + count
                break
              case 'Min number':
                let min: number = Number.MAX_VALUE
                for (let val in values) {
                  let valNum: number = parseInt(val, 10)
                  if (valNum < min) min = valNum
                }
                if (min !== Number.MAX_VALUE)
                  valueString = '' + min
                break
              case 'Max number':
                let max: number = Number.MIN_VALUE
                for (let val in values) {
                  let valNum: number = parseInt(val, 10)
                  if (valNum > max) max = valNum
                }
                if (max !== Number.MIN_VALUE)
                  valueString = '' + max
                break
              case 'Number range':
                let rmin: number = Number.MAX_VALUE
                let rmax: number = Number.MIN_VALUE
                for (let val in values) {
                  let valNum: number = parseInt(val, 10)
                  if (valNum > rmax) rmax = valNum
                  if (valNum < rmin) rmin = valNum
                }
                if (rmax !== Number.MIN_VALUE || rmin !== Number.MAX_VALUE)
                  valueString = '' + (rmin !== Number.MAX_VALUE ? rmin : '') + '-' + (rmax !== Number.MIN_VALUE ? rmax : '')
                break
                case 'Min date':
                  let dmin: Date = new Date(8640000000000000)
                  for (let val in values) {
                    let valNum: Date = new Date(val)
                    if (valNum < dmin) dmin = valNum
                  }
                  if (dmin.getTime() !== 8640000000000000)
                    valueString = '' + dmin.toISOString().substring(0, 10)
                  break
                case 'Max date':
                  let dmax: Date = new Date(-8640000000000000)
                  for (let val in values) {
                    let valNum: Date = new Date(val)
                    if (valNum > dmax) dmax = valNum
                  }
                  if (dmax.getTime() !== -8640000000000000)
                    valueString = '' + dmax.toISOString().substring(0, 10)
                  break
                case 'Date range':
                  let drmin: Date = new Date(8640000000000000)
                  let drmax: Date = new Date(-8640000000000000)
                  for (let val in values) {
                    let valNum: Date = new Date(val)
                    if (valNum > drmax) drmax = valNum
                    if (valNum < drmin) drmin = valNum
                  }
                  if (drmax.getTime() !== -8640000000000000 || drmin.getTime() !== 8640000000000000)
                    valueString = '' + (drmin.getTime() !== 8640000000000000 ? drmin.toISOString().substring(0, 10) : '') + '-' + (drmax.getTime() !== -8640000000000000 ? drmax.toISOString().substring(0, 10) : '')
                  break
              default: console.log('unknown type: ' + heading.type)
            }
            row.push(valueString)
          })
        }
      }
      $scope.deleteOutputColumn = (column: string) => {
        delete $scope.output.mappings[column]
        updateOutput()
      }
      $scope.deleteOutputMapping = (rowIndex: number) => {
        $scope.output.headings.splice(rowIndex, 1)
        for (let key in $scope.output.mappings) $scope.output.mappings[key][rowIndex] = undefined
        updateOutput()
      }
      $scope.addMapping = (key: string, column: number, origColumn: string) => {
        if (!origColumn) return
        if (!$scope.output.mappings[key][column]) $scope.output.mappings[key][column] = []
        $scope.output.mappings[key][column].push(parseInt(origColumn, 10))
        updateOutput()
      }
      $scope.deleteMapping = (key: string, column: number, origColumnIndex: number) => {
        $scope.output.mappings[key][column].splice(origColumnIndex, 1)
        updateOutput()
      }
      $scope.addMappingKey = (selectedHeading: string) => {
          $scope.output.mappings[selectedHeading] = []
          updateOutput()
      }
      $scope.addHeading = () => {
        $scope.output.headings.push(new Heading())
        updateOutput()
      }
      $scope.selectMappingToLoad = () => {
        $uibModal.open({
          templateUrl: 'selectMapping',
          scope: $scope
        })
      }
      $scope.loadMapping = (mappings: Table) => {
        $scope.output.headings = mappings.headings
        $scope.output.mappings = mappings.mappings
        updateOutput()
      }
      $scope.showSaveMappingDialog = () => {
        $uibModal.open({
          templateUrl: 'saveMapping',
          scope: $scope
        })
      }
      $scope.saveMapping = (name: string) => {
        $localStorage.mappings[name] = {
          mappings: $scope.output.mappings,
          headings: $scope.output.headings
        }
      }
      $scope.deleteSavedMapping = (name: string) => {
        delete $localStorage.mappings[name]
      }
      $scope.importMapping = (file: File) => {
        if (!file) return
        let reader: FileReader = new FileReader()
        reader.onload = () => {
          let mappings: Table = JSON.parse(reader.result)
          $scope.output.mappings = mappings.mappings
          $scope.output.headings = mappings.headings
          updateOutput()
          $scope.$digest()
        }
        reader.readAsText(file)
      }
      $scope.exportMapping = () => {
        let mappings: Table = new Table()
        mappings.mappings = $scope.output.mappings
        mappings.headings = $scope.output.headings
        saveAs(new Blob([JSON.stringify(mappings)], {type: 'application/json'}), 'mappings.json')
      }
      $scope.output = new Table()
      let handleError: (error: IError) => void = (error: IError) => $scope.error = error
      let init: (data: string[][]) => void = (data: string[][]) => {
        $scope.originalHeadings =  data[0].map((column, index) => $scope.firstRowIsHeader ? column : 'Column ' + (index + 1))
        $scope.originalData = data
        $scope.originalData.splice(0,  ($scope.skipRows ? $scope.skipRows : 0) + ($scope.firstRowIsHeader ? 1 : 0))
      }
      $scope.firstRowIsHeader = true
      $scope.loadFile = (file: File) => {
        if (!file) return
        $scope.fileName = file.name
        if (file.type === 'text/csv')
          Papa.parse(file, { complete: (csv): void => {
            if (csv.errors.length !== 0)
              handleError({data: csv.errors.map(e => e.message).join('\n')})
            else {
              $scope.sheets = [ new Sheet('Table') ]
              $scope.firstRowIsHeader = true
              $scope.sheets[0].data = csv.data
              $scope.sheets[0].active = true
              $uibModal.open({
                templateUrl: 'selectSheet',
                scope: $scope,
                size: 'lg'
              })
            }
          }
          , error: (error: PapaParse.ParseError): void => handleError({data: error.message})})
        else {
          let reader: FileReader = new FileReader()
          reader.onload = () => {
            try {
              let workBook: any = XLSX.read(reader.result, {type: 'binary'})
              $scope.sheets = []
              workBook.SheetNames.forEach(sn => {
                let sheet: Sheet = new Sheet(sn)
                let sheetJson: {}[] = XLSX.utils.sheet_to_json(workBook.Sheets[sn])
                sheet.data = [[]]
                for (let header in sheetJson[0]) if (header.indexOf('__') !== 0) sheet.data[0].push(header)
                XLSX.utils.sheet_to_json(workBook.Sheets[sn]).forEach(row => {
                  let row2: string[] = []
                  sheet.data[0].forEach(h => row2.push(row[h]))
                  sheet.data.push(row2)
                })
                $scope.sheets.push(sheet)
              })
              $scope.firstRowIsHeader = true
              $scope.sheets[0].active = true
              $uibModal.open({
                templateUrl: 'selectSheet',
                scope: $scope,
                size: 'lg'
              })
            } catch (err) {
              handleError({data: err.message})
            }
          }
          reader.readAsBinaryString(file)
        }
      }
      $scope.loadSheet = () => {
        $scope.sheets.filter(s => s.active).forEach(s => {
          init(s.data)
        })
      }
      $scope.saveCSVFile = () => {
        let data: string[][] = [$scope.output.headings.map(h => h.label)].concat($scope.output.data)
        saveAs(new Blob([Papa.unparse(data)], {type: 'text/csv'}), 'mare-' + $scope.fileName + ($scope.fileName.indexOf('.csv', $scope.fileName.length - 4) !== -1 ? '' : '.csv'))
      }
    }
  }
}

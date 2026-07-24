[English](CONTRIBUTING.md) · **한국어**

# 기여 가이드

이 문서는 저장소 구조·빌드·테스트·코드 규칙을 정리한 개발자용 안내다.  
라이브러리 사용법은 [README_ko.md](README_ko.md)와 [docs/reference_ko.md](docs/reference_ko.md)를 참고한다.

---

## 저장소 구조

Maven 멀티모듈 프로젝트다.

| 모듈 | 대상 | 컴파일 |
|---|---|---|
| `pxl-javax` | Java 8+, `javax.*` | source/target 8 |
| `pxl-jakarta` | Java 17+, `jakarta.*` | source/target 17 |

> ⚠️ 소스는 `pxl-javax`에만 있다. `pxl-jakarta`는 직접 편집하지 않는다.

`pxl-jakarta`에는 직접 작성한 소스가 없다.
빌드 시 형제 모듈 `pxl-javax`의 소스 트리(`src/main/java`·`src/main/resources`, test는 `src/test`)를 파일시스템 복사해 아래 문자열 치환을 거친 뒤 컴파일해 생성된다(`pxl-jakarta/pom.xml`의 `maven-antrun-plugin`).

```
import javax.annotation.   →   import jakarta.annotation.
import javax.validation.   →   import jakarta.validation.
```

따라서:

- 코드 수정은 항상 `pxl-javax`에서 한다. `pxl-jakarta` 쪽 동작을 바꾸려면 `pxl-javax`를 고쳐 다시 빌드한다.
- `pxl-javax`의 이식 대상 import는 정확히 `javax.annotation.*` / `javax.validation.*` 로만 제한한다 — 그 외 `javax.*`(예: `javax.imageio`)는 치환되지 않으므로 jakarta 변형에서 그대로 남는다.

베이스 패키지는 `io.github.hclimkr.pxl`이며, 공개 API(`Pxl`·`annotation`·`option`·`exception`·`styler`·`util`·`constraint`·`builder`)와 내부 구현(`internal/*`)의 경계는 `.internal.` 네이밍 관례로 유지한다(사용자는 `internal.*`를 참조하지 않는다).

---

## 개발 환경

- JDK 17이 PATH에 있어야 한다(`pxl-javax`는 8로 컴파일하지만 빌드 자체는 JDK 17에서 수행).
- Lombok이 전반에 쓰이며 각 모듈의 `maven-compiler-plugin`에 애노테이션 프로세서로 연결돼 있다. IDE에서 Lombok 플러그인을 활성화한다.

---

## 빌드 & 테스트

`pxl-jakarta`는 `pxl-javax`의 소스 트리를 직접 복사하므로 아티팩트 install에 의존하지 않지만(reactor 순서는 루트 `<modules>` 선언 순서로 유지된다),
두 모듈을 한 번에 빌드하려면 루트에서 실행하기를 권장한다.

```bash
mvn clean install                                    # 두 모듈 빌드 + 테스트
mvn -pl pxl-javax test                               # javax 모듈 테스트 (모든 로직이 여기 있음)
mvn -pl pxl-javax test -Dtest=PxlExcelExportTests    # 단일 테스트 클래스
mvn -pl pxl-javax test -Dtest=PxlExcelImportTests#someMethod   # 단일 테스트 메서드
mvn clean install -DexcludedGroups=network           # 네트워크 의존 테스트 제외
mvn install -DskipTests                              # 테스트 없이 빌드
```

- 테스트는 JUnit 5(Jupiter) + AssertJ 기반이다.
- 외부 이미지(picsum.photos)에 접근하는 테스트는 `@Tag("network")`이며 `-DexcludedGroups=network`로 제외할 수 있다.
- export 결과 파일은 `pxl-javax/target/test-outputs/*.xlsx`(및 `*.xls`)에 기록되며(`TestPaths.EXPORT_DIR`), `target/` 하위이므로 `mvn clean` 시 삭제된다.

---

## 테스트 작성 규칙

- 새 `*Tests.java`를 만들지 않는다. 실행 클래스는 루트 테스트 패키지 `io.github.hclimkr.pxl`에 기능별 카테고리로 분류돼 있으니(`PxlRoundTripTests`, `PxlTypeConversionTests`, `PxlExcelImportTests`, `PxlExcelExportTests`, `PxlCsvImportTests`, `PxlNameMatchingTests`, `Pxl{Column,Sheet,Workbook}OptionTests`, `PxlStyleTests`, `PxlSampleExcelExportTests`, `PxlI18nTests`, `PxlExceptionTests`, `PxlRegressionTests`, `Pxl{Picture,PictureNetwork}ExportTests`, `PxlLargeDataTests`, `PxlUtilityTests`) 해당 기능 클래스에 메서드를 추가한다.
- 명명 규칙: 클래스명 = 기능 카테고리, 메서드명 = `기능_상황_결과` camelCase (예: `importExcel_mergedRegion_streaming_throws`).
- 바인딩 대상이 되는 애노테이션 붙은 DTO 픽스처는 `tcdata/` 하위 패키지에 둔다. 입력 데이터는 대부분 각 테스트에서 코드로 직접 생성한다.
- 로직·성능을 바꾸는 수정에는 회귀/특성화 테스트를 함께 추가한다.

---

## 코드 규칙

- JavaDoc(`/** */`)과 인라인·블록 주석(`//`·`/* */`)은 영어로 작성한다.  
  단 예외·진단 메시지 i18n 번들(`pxl-messages{,_ko}.properties`, base=영어)에 키로 두고 `PxlI18nDiagnostic.get(...)`으로 해석하도록 한다.
- 공개 API의 필수 인자는 파라미터 애노테이션이 아니라 `PxlAssertSupport`(`notNull`/`notEmpty`/`notBlank`) 명령형 호출로 검증한다.  
  nullability 표기는 내부 `@Nullable`(`internal/constraint`)을 쓴다.
- 새 필드 타입 지원을 추가하려면 `internal/codec/`에 codec 클래스를 추가하고 `PxlCellResolver`의 세 메서드(`parseDataValueFromCell`·`parseDataValueFromString`·`buildDataCell`) 모두에 분기를 추가한다.
- 작업 트리는 CRLF 줄바꿈을 쓰지만 `.gitattributes`가 없어, `core.autocrlf=true`(이 저장소 기준) 설정에서 Git이 커밋되는 blob을 LF로 정규화한다. 대량 치환 후 EOL이 섞이지 않도록 주의한다.

---

## Pull Request

1. `main`에서 브랜치를 딴다.
2. 변경에 대응하는 테스트를 추가/수정한다.
3. `mvn clean install -DexcludedGroups=network`가 통과하는지 로컬에서 확인한다.
4. PR을 올린다.

---

## 라이선스

기여한 코드는 프로젝트와 동일하게 [Apache License 2.0](LICENSE) 하에 배포되는 데 동의하는 것으로 간주한다.

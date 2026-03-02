# Changelog

## [2.0.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.7.0...v2.0.0) (2026-03-02)


### ⚠ BREAKING CHANGES

* migrate ci workflow from poetry to uv ([#123](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/123))

### Features

* **deps:** update dependency jgm/pandoc to v3.9 ([4e5f444](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/4e5f444c7ddadeb06da25e27ecd570d11b0db066))
* migrate ci workflow from poetry to uv ([#123](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/123)) ([72f4d94](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/72f4d94027b8bfb4c403e6ea5d7198c4d118e361))
* migrate dockerfile from poetry to uv ([#124](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/124)) ([5d1748a](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/5d1748a1f2a78e03e9390423de3076367251f688))
* switch to python 3.14 ([#127](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/127)) ([e77c55d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/e77c55dbbb32e8604dca542a5b35e94cddfdf61e))


### Bug Fixes

* complete poetry-to-uv migration for CI and local tooling ([b69cfae](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/b69cfae37a9dd974ccec919ecb96e9729982af1e))
* **deps:** update dependency fastapi to v0.129.0 ([#125](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/125)) ([557a056](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/557a056efe321be6d6e00103038d51c89a76f13b))
* **deps:** update dependency fastapi to v0.129.2 ([d809eaf](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d809eaf95c2c81bb95a4b28afb99f6af4e1dd85b))
* **deps:** update dependency fastapi to v0.131.0 ([a0313f8](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/a0313f80c95b24a20ee771e5bab6b2bf1706903b))
* **deps:** update dependency fastapi to v0.132.0 ([5a21382](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/5a213829b3f738e75e187522e05c8bda62420787))
* **deps:** update dependency fastapi to v0.133.0 ([87ccccd](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/87ccccdb83b5e7008c0c454f74c487e63a700e37))
* **deps:** update dependency fastapi to v0.133.1 ([cce45f7](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/cce45f76b1b628afdef33fb75313933d16903c74))
* **deps:** update dependency uvicorn to v0.41.0 ([566071d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/566071d05c2f732f8c7f86bd8736736d7e07db47))
* remove redundant response_model ([#130](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/130)) ([f01a765](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/f01a765ff07409ac7993076df5b70a5671968c79))
* update tox envlist to py314 after Python 3.14 migration ([#128](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/128)) ([c400640](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/c400640386eb8af5db1e7fe277c59325f2466d9a))

## [1.7.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.6.0...v1.7.0) (2026-01-22)


### Features

* Replace generated hyperlinks for ToC, ToF, and ToT with macros ([#114](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/114)) ([8bc34b3](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/8bc34b39098d2a71aae103e5c914058b37a2497f)), closes [#113](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/113)


### Bug Fixes

* headers/footers from template are missing in all document segments ([#116](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/116)) ([d36dea2](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d36dea229259166205c808c9367946c301d715a5)), closes [#115](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/115)
* increase multipart form data size limit ([#120](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/120)) ([256f912](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/256f9120c6ffc9e564d1198ed2b5ff87b9d1d115)), closes [#119](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/119)

## [1.6.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.5.2...v1.6.0) (2026-01-13)


### Features

* add PPTX and template endpoint tests to container integration tests ([#110](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/110)) ([97a59c6](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/97a59c6785f89f761f42d5449cd0071d5b6f1eaa)), closes [#109](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/109)
* add PPTX conversion endpoint with template support ([#86](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/86)) ([9c35958](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/9c35958665ee18ea4196f6a0a415ef8020203dc4))
* migrate pyproject.toml to PEP 621/735 standard format ([#108](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/108)) ([771d99a](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/771d99adb2669edcf10e5e49c00c47ab1763827c)), closes [#100](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/100)


### Bug Fixes

* **deps:** update dependency fastapi to v0.124.2 ([5bc9607](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/5bc9607a82c69bc58e94b9f9599d81482ee80ebe))
* **deps:** update dependency fastapi to v0.124.4 ([52c27ba](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/52c27ba1d35095de4c8ffdcb5b34eecd6787cdd1))
* **deps:** update dependency fastapi to v0.125.0 ([d1a97ef](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d1a97ef570aa73a35c5386216cc9e0db2a366f57))
* **deps:** update dependency fastapi to v0.126.0 ([6ab25ba](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/6ab25ba73f404c30e785af23276300d120427624))
* **deps:** update dependency fastapi to v0.127.0 ([034106c](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/034106c0e541efb1bf0c975f8aeca91603a59194))
* **deps:** update dependency fastapi to v0.127.1 ([28d21b1](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/28d21b177a8800aae64e945c832f9eb1b72a7b08))
* **deps:** update dependency fastapi to v0.128.0 ([601f0ed](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/601f0edda525e70b4937263da165ac361d4869c7))
* **deps:** update dependency python-multipart to v0.0.21 ([c414c16](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/c414c16a532820ea4bcbbcba899707524aebe406))
* **deps:** update dependency uvicorn to v0.40.0 ([b01a603](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/b01a60342e09a9cce7f57cfd1ec8b834e519d644))
* remove max_part_size from pptx endpoint and improve test quality ([#107](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/107)) ([0d4109e](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/0d4109e5ac4d695020806e87d8c05214206ecc5b))
* resolve SonarCloud quality gate failures ([#96](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/96)) ([a446fb1](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/a446fb1c007360ec08a961772748356cf09bcd84)), closes [#95](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/95)
* resolve SonarCloud quality gate failures ([#99](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/99)) ([4028fc8](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/4028fc83f707b2be319a4614b0a0749089ca6b57))

## [1.5.2](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.5.1...v1.5.2) (2025-12-10)


### Bug Fixes

* patch XML parser to handle large documents ([#92](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/92)) ([d5f9b0e](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d5f9b0e22167e76286ddd386b7ee601ba8f9f82b)), closes [#91](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/91)

## [1.5.1](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.5.0...v1.5.1) (2025-12-08)


### Bug Fixes

* **deps:** update dependency asgiref to v3.11.0 ([29b1c20](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/29b1c208293ec3adb40f95f72f3437eb76c6b638))
* **deps:** update dependency fastapi to v0.121.3 ([1ddadaa](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/1ddadaa1b10a74688a8e6b15f8a1a747a3e71b6f))
* **deps:** update dependency fastapi to v0.122.0 ([e4e6991](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/e4e69917813ba5df30e4aca464253e055006c3fa))
* **deps:** update dependency fastapi to v0.123.0 ([3ade322](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/3ade322e50fd49812db82f4dc74fa172e12d2ead))
* **deps:** update dependency fastapi to v0.123.10 ([8ddd56c](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/8ddd56cbb970dadd844e7158cb9d8abe27f7804b))
* **deps:** update dependency fastapi to v0.123.4 ([25de981](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/25de9813e486100b832979c4d613935f1b0c3cf0))
* **deps:** update dependency fastapi to v0.123.5 ([9d36ac7](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/9d36ac70e949db566b5d05af85933ed271a93bbc))
* **deps:** update dependency fastapi to v0.123.7 ([2246d51](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/2246d5163eebab9b90168b0123567cf590f220ed))
* **deps:** update dependency fastapi to v0.123.9 ([da58329](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/da58329877a6f309b4714889c89a47134992fcb7))
* **deps:** update dependency fastapi to v0.124.0 ([3980c4d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/3980c4d33332e43bd9f02a1f8aa5468c9f406825))
* **deps:** update dependency jgm/pandoc to v3.8.3 ([261f875](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/261f8751b3d3188558179d206b3b34eee0cce539))
* patch XML parser to handle large documents ([#90](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/90)) ([e71bd8d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/e71bd8d425e02e8a9825e6f7823d7c29fea8798d)), closes [#89](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/89)

## [1.5.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.4.0...v1.5.0) (2025-11-17)


### Features

* export landscape page ([#85](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/85)) ([6a06b40](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/6a06b40b97bb547e042f87adcbdbc50dabd994e6)), closes [#84](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/84)


### Bug Fixes

* **deps:** update dependency fastapi to v0.121.1 ([db1473b](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/db1473be6159594db50e3c0fee39444e9e1a13b8))
* **deps:** update dependency fastapi to v0.121.2 ([b914095](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/b91409535294f8507d2ba0049dad6e220e6197c7))

## [1.4.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.3.0...v1.4.0) (2025-11-04)


### Features

* ability to set custom page size and orientation for docx files ([#80](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/80)) ([008e892](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/008e892cd3a9d1a9830843938c95948c61777628)), closes [#79](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/79)
* make memory limit configurable ([#78](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/78)) ([95450d5](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/95450d5c3ab05fac0bfe85d1be9bbf45e7d9bb08)), closes [#75](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/75)


### Bug Fixes

* **deps:** update dependency fastapi to v0.118.3 ([f7fcd11](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/f7fcd1165189073b6c9f256a4d8802c4c669bb6a))
* **deps:** update dependency fastapi to v0.119.0 ([3536dd2](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/3536dd233587ac57a05309b5ef18b0c9d0b219e3))
* **deps:** update dependency fastapi to v0.119.1 ([498d52d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/498d52de679bf750a3cbc4007e5f31917819b9fb))
* **deps:** update dependency fastapi to v0.120.0 ([1ee0d2f](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/1ee0d2f85a24a1773b83ee6ba578b4dd2ca8cb71))
* **deps:** update dependency fastapi to v0.120.1 ([71a8592](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/71a8592cfce7daf7fa9fa9b460998a65d19f42f6))
* **deps:** update dependency fastapi to v0.120.2 ([d923c23](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d923c23771a93ceb5c222d98de46512a42d3b5ab))
* **deps:** update dependency fastapi to v0.120.3 ([db752b2](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/db752b23221e7546f24473373e6ef82b3b5f7925))
* **deps:** update dependency fastapi to v0.120.4 ([9f3e1d9](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/9f3e1d9aa828222ae0579a1e5fbe2de8607b1d79))
* **deps:** update dependency fastapi to v0.121.0 ([66c2abd](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/66c2abddb759c54107b72bb8b70fb1eba9790ba0))
* **deps:** update dependency jgm/pandoc to v3.8.2.1 ([6a88d67](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/6a88d67a678f5bf214f5e7dd4f13c8777f269833))
* **deps:** update dependency starlette to ^0.49.0 [security] ([61b5860](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/61b5860544470dbd69e1ad53896e9989bb507173))
* **deps:** update dependency starlette to v0.49.1 [security] ([e0a9b56](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/e0a9b56644898be4cd6f35ffe624cecf8cc5a442))
* **deps:** update dependency uvicorn to v0.38.0 ([3c28713](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/3c28713681462160e43667c3d3ff903d05537cbd))

## [1.3.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.2.0...v1.3.0) (2025-10-09)


### Features

* add option to generate toc in pandoc ([#72](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/72)) ([cc8407f](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/cc8407fab9a8ae34617eaaa8d6c78df1037bc84d)), closes [#71](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/71)


### Bug Fixes

* **deps:** update dependency asgiref to v3.10.0 ([3a4be1d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/3a4be1d64e0de85634eac275cb216177288fcbf9))
* **deps:** update dependency asgiref to v3.9.2 ([be8ee05](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/be8ee056d45a9de0e55816af673eb9f949aaa89b))
* **deps:** update dependency fastapi to v0.116.2 ([10692ab](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/10692ab817a7fc0cb168edc3db9192bd1a66651a))
* **deps:** update dependency fastapi to v0.117.1 ([b4b1ccb](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/b4b1ccb99d81b5dca7d96628d8bbd3f314296fda))
* **deps:** update dependency fastapi to v0.118.0 ([#69](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/69)) ([686be0e](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/686be0ef2ffa7242be423fb94d6a85f1b0d9e635))
* **deps:** update dependency fastapi to v0.118.1 ([90bbbbd](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/90bbbbd6bad2d45cc2e9f540b8f49e65085eeac7))
* **deps:** update dependency fastapi to v0.118.2 ([94eeef2](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/94eeef2a5a604bce3023808a0452465952495e70))
* **deps:** update dependency jgm/pandoc to v3.8.1 ([f6ac063](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/f6ac063605cb339d6d805ab897e3675f38ed9c56))
* **deps:** update dependency jgm/pandoc to v3.8.2 ([a90b840](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/a90b840173758a874c874e2366048f3364293660))
* **deps:** update dependency starlette to ^0.48.0 ([#66](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/66)) ([a2a2490](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/a2a249014fee33c5ac2f2bd8576c839a7251f9c2))
* **deps:** update dependency uvicorn to v0.36.0 ([815f3ea](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/815f3eadac3cb4deeebee2c59185030cba5d5346))
* **deps:** update dependency uvicorn to v0.37.0 ([e84f0a4](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/e84f0a40701bc349fd4a044aa03dca87b5859e5f))

## [1.2.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.1.0...v1.2.0) (2025-09-11)


### Features

* **deps:** update dependency jgm/pandoc to v3.8 ([fb446b0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/fb446b0272dd4e70162506f10fa338504782e1af))


### Bug Fixes

* **deps:** update dependency starlette to v0.47.3 ([b5e384f](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/b5e384f7649e9549a662d78d26ed34732a4efd82))
* set max_part_size for form data requests ([#65](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/65)) ([83659c3](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/83659c3363ce8172d3ef4ce30ca47290d8b38fd4)), closes [#64](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/64)

## [1.1.0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/compare/v1.0.0...v1.1.0) (2025-07-30)


### Features

* add PDF generation ([#34](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/34)) ([2666ee0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/2666ee0d8d5589fd169923fa4d0d1db8bac4d5f4)), closes [#33](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/33)
* **deps:** update dependency jgm/pandoc to v3.7 ([9a16aa6](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/9a16aa6005b900281620c8c730848254ab955c8d))
* replace tini with --init ([#32](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/32)) ([2ada8a1](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/2ada8a1ef65f7ee83034ff9217b60ffa3b77cb23))
* use fastapi instead of flask ([#30](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/30)) ([f32ebe5](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/f32ebe5400d967dc90b9ea328ca5d1681dc88200))


### Bug Fixes

* **deps:** update dependency asgiref to v3.9.0 ([#53](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/53)) ([c18a225](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/c18a22513fcd6fbace2ec1ebb7a41eb037869add))
* **deps:** update dependency asgiref to v3.9.1 ([db9c4ba](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/db9c4ba8c53dcce66a5c11998da9cc2300079c64))
* **deps:** update dependency fastapi to v0.115.14 ([#44](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/44)) ([4bd6f92](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/4bd6f9205b374c5fb8abfba0104877a6cae7e8ae))
* **deps:** update dependency fastapi to v0.116.0 ([9646565](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/96465656eedb1b0a80f4677936df211b459361ab))
* **deps:** update dependency fastapi to v0.116.1 ([0304ae6](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/0304ae6a05d7470004a31bf51c158b75b2eea973))
* **deps:** update dependency python-docx to v1.2.0 ([#37](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/37)) ([d7ba30f](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/d7ba30f2b36a55e0a40c7db26bd9483248aa09fa))
* **deps:** update dependency uvicorn to v0.34.3 ([#36](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/36)) ([fc24c2d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/fc24c2d504cee2153446c139dcd4ca71d17e35e1))
* **deps:** update dependency uvicorn to v0.35.0 ([#50](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/50)) ([122b042](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/122b042cacf621fbdd43cc310b0573db647fd233))
* **deps:** update versioning template for pandoc in renovate.json ([#41](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/41)) ([26205d0](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/26205d01cfe68cdb7669a325c2ec5eee3e3c8c2c))
* remove redundant imports in test_docx_post_process.py ([#51](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/51)) ([60a31cb](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/60a31cbb05adf11bbcab398f39882cf14f951725))

## 1.0.0 (2025-04-10)


### Features

* add tini into dockerfile ([#20](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/20)) ([9a8cb62](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/9a8cb62de7674bbb04ea26662775498721fe50ec))
* Initial implementation of pandoc-service ([#5](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/5)) ([2eab04d](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/2eab04db8288e8d47f39ebc958f147b708d8f46d))
* logging in docker container for Splunk ([#15](https://github.com/SchweizerischeBundesbahnen/pandoc-service/issues/15)) ([694fac4](https://github.com/SchweizerischeBundesbahnen/pandoc-service/commit/694fac4004deb2526c3498fadce23e2fba43c54a))

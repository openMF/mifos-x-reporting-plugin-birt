/**
 * Copyright since 2026 Mifos Initiative
 *
 * <p>This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0. If a copy
 * of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
package org.apache.fineract.infrastructure.report.service;

import static org.apache.fineract.infrastructure.core.domain.FineractPlatformTenantConnection.toJdbcUrl;
import static org.apache.fineract.infrastructure.core.domain.FineractPlatformTenantConnection.toProtocol;

import jakarta.ws.rs.core.MultivaluedMap;
import jakarta.ws.rs.core.Response;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import javax.sql.DataSource;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.fineract.infrastructure.core.api.ApiParameterHelper;
import org.apache.fineract.infrastructure.core.config.FineractProperties;
import org.apache.fineract.infrastructure.core.domain.FineractPlatformTenant;
import org.apache.fineract.infrastructure.core.domain.FineractPlatformTenantConnection;
import org.apache.fineract.infrastructure.core.exception.PlatformDataIntegrityException;
import org.apache.fineract.infrastructure.core.service.ThreadLocalContextUtil;
import org.apache.fineract.infrastructure.core.service.database.DatabasePasswordEncryptor;
import org.apache.fineract.infrastructure.dataqueries.data.ReportExportType;
import org.apache.fineract.infrastructure.report.annotation.ReportService;
import org.apache.fineract.infrastructure.security.constants.TenantConstants;
import org.apache.fineract.infrastructure.security.service.PlatformSecurityContext;
import org.eclipse.birt.report.engine.api.EXCELRenderOption;
import org.eclipse.birt.report.engine.api.HTMLRenderOption;
import org.eclipse.birt.report.engine.api.HTMLServerImageHandler;
import org.eclipse.birt.report.engine.api.IEngineTask;
import org.eclipse.birt.report.engine.api.IGetParameterDefinitionTask;
import org.eclipse.birt.report.engine.api.IImage;
import org.eclipse.birt.report.engine.api.IPDFRenderOption;
import org.eclipse.birt.report.engine.api.IParameterDefn;
import org.eclipse.birt.report.engine.api.IReportEngine;
import org.eclipse.birt.report.engine.api.IReportRunnable;
import org.eclipse.birt.report.engine.api.IRunAndRenderTask;
import org.eclipse.birt.report.engine.api.PDFRenderOption;
import org.eclipse.birt.report.engine.api.RenderOption;
import org.eclipse.birt.report.model.api.DesignElementHandle;
import org.eclipse.birt.report.model.api.EmbeddedImageHandle;
import org.eclipse.birt.report.model.api.LibraryHandle;
import org.eclipse.birt.report.model.api.OdaDataSourceHandle;
import org.eclipse.birt.report.model.api.ReportDesignHandle;
import org.eclipse.birt.report.model.api.SlotHandle;
import org.eclipse.birt.report.model.api.StructureFactory;
import org.eclipse.birt.report.model.api.activity.SemanticException;
import org.eclipse.birt.report.model.api.elements.DesignChoiceConstants;
import org.eclipse.birt.report.model.api.elements.structures.EmbeddedImage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.ApplicationContext;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;

@Service
@ReportService(type = "BIRT")
public class BirtReportingProcessServiceImpl implements ReportingProcessService {

  private static final Logger logger = LoggerFactory.getLogger(ReportingProcessService.class);

  private final String mifosBaseDir = System.getProperty("user.home") + File.separator + ".mifosx";
  private final DatabasePasswordEncryptor databasePasswordEncryptor;

  @Value("${FINERACT_BIRT_REPORTS_PATH:}")
  private String fineractBirtBaseDir;

  @Value("${FINERACT_BIRT_REPORTS_LOCALE:}")
  private String fineractBirtLocale;

  private final IReportEngine reportEngine;
  private final DataSource tenantDataSource;

  private final FineractProperties fineractProperties;

  private final ApplicationContext applicationContext;
  private final PlatformSecurityContext context;
  private final ApplicationContext contextVar;

  // Centralized Base64 logo for consistent rendering across environments and easier maintenance
  private static final String CENTRAL_LOGO_BASE64 =
      """
  iVBORw0KGgoAAAANSUhEUgAAAgAAAAEACAYAAADFkM5nAAA6DElEQVR42u3deZwcZZ0/8M+3emYSjhwz
                Uz3T0zPgyKVyeAD6IwGEZEHlVFQit4ACq6sQbljkSgigCEncxXMRZRdX8UJJRDkSBAmyS4IuBtYlQIDJ
                XN0zgQRIZqa7vr8/eo6eme6q6que6snn/dIXyXRVPU9VZub51vN8n+cBiIiIiIiIiIiIiIiIiIiIiIiI
                iIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiI
                iIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiI
                iIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiI
                iIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIQkdMV4CIiMa7dv3BHwKcD5uuB/kTEXnm
                hn3XrjNdj0LVmK4AERGN50i633LkdgC7mq4LeXOgZwCougDAMl0BIiIab8m+z77qiBwD4C3TdSFXqpDr
                Fu277iemK1IMDgEQEYXUNc99cK4l1u8BzDBdF5pEVeXimw5Yu9x0RYrFHgAiopBacsBf1gDWMQC2mq4L
                jVP1jT/AAICIKNQW7//MkwwCwkWhV1d74w9wCICIqCpc+7eDDwWcB8HhAKMUetVN+z/7ddP1KAcGAERE
                VYJBgFlTqfEHGAAQEVUVBgFmTLXGH2AAQERUdRgEBGsqNv4AAwAioqrEICAYU7XxBxgAEBFVLQYBlTWV
                G3+AAQARUVVjEFAZU73xB7gOABFRVeM6ARUgcvVUb/wBBgBERFVv8f7PPOmIHAsGAaUTuXrxfmtvDaKo
                1ptfaARgm7pVBgBERJWycM25WLjm1CCKWrLf2j8xCChRgI3/7rf8T701pL/ffdH6ZlO3ywCAiKhyDgXw
                E1z81D8HURiDgBIE3Pg7g7UPqcjBJm+ZAQARUaWpLsHFa64KoqisIIBbCfsVdOOfqn0YFow2/gADACKi
                YChuCTgIOAYMArwF3vjXPQzgoEzZZm+dAQARUVAYBIRLgI1/7OYXo+l03WMqOEghUNOtPxgAEBEFi0FA
                OATc+EeQfhTA+zNlZ/5vOghgAEBEFDTFLVi45uogimIQkEPAjb8F51FVOUBlQoPPIQAioh3SzQwCDDDQ
                +ENxwMjXVARqvvcfAAMAIiKTGAQEKfDGH48qcMBIl39WRUIRCDAAoILE43HbjtnzorHoAjtmH9He3j7d
                ZH3q6+tnxWKxdrvV3qe5uXmP2e2zZ5t+RkQFujngdQJ2zCAgwMa/+bYNTZbgURU9AMMJfwogVyAwiDpz
                j8RYyVRVmpub3522dDGgpwCIZH20VYB/3WXaLos2bty4vdL1aGtr22nAGThVVU+EyEegaMlx2Bsi+F9V
                eU6g6x3Hea4GNU/39PS8bfo50g5m4Zq7AJzr61iRa7B0zs1BVOva9QceDsXvAOxq8OkERkX++ab91t4S
                RFnNt21okpTzqAD7j6/E2B8k62tO2tq/+7q915t4LuwBIE923D4wbTl/BvR0jG/8AWCGAle/NfD24/X1
                9bMqWA1pjDV9eSA98Loq7gLkk3kafwCYrYpDAD1PgWViWY+mLScZjUUXmH6WRHmpLgmqJ2DxfuuegGCH
                WCwo8Mbf0Uchsr+KjLX5QNbbf9bXOQRAYWbb9gxAVgBo8jj0wzXTa35YiTq0tbXtZMea7hfROxVoLPIy
                01VwQyWfFVHJVJdg4VPXBFHUjhAEGGn8NevNP9c4/2gQMBwImBsBqJ4AIBqPHmbHmn7TFG861HRdwqC9
                vX263dL0PTtmX1rJcrRWrnR50x5/LPDphtamOWWuQmQgPXAvRE8sw7XeVf4nRFRuehODgNIF3fjDwSoF
                9lc/CX9ZvQEmVU0AoI5cDtETHdUpv0ezH28NvHM2oOdD8NVKliOQkwo63tHPlbN8O9Z0sQIF1YGo+jEI
                KEWQjX/TkpeboVgFwX4jLXu+hL+cwwIGVU0AANHZw3+a09zc/G7T1TGsFtArAAAiFetAam5u3gXQ9xVy
                jgUUdLybmW0zGyAo3xQpwVAZHw9RhelNuOjJrwVR0lQKAgJt/Je93Ixa51EF9gMwYZw/KxDIJjkCAUOq
                JwAA2kbqnLY0kHmzYWXH7c8DyARBilg0Gq1IJm9PT88gCoxRHZTv+7rWmf5FQBvKdkNa/b/caAcjsphB
                gH+BN/5Dugoi+8E14W+4+z9nIGD2eVVLAFALYPfRv4meHYvF2k1XypAIVK7I+rugDntVqKwhKDoLPOd/
                y1W4qJ7s47C0CP4M6HchshjQi6FyhQC3quBHIvgzBG8CgAK/rNBzIqocBgG+BN74p3QVgH1HvyhuXfyZ
                r+fKDzCpxmjpPtmt9rvhZNVVUZuS1JUAvmS6bkGLxqNnqWLvcV9MYx8Af6lIgYL7ACz0e7gF/KocxcZi
                sWgK6YPcj9JX1Yl8LtnT87TX9fbaa69pGzZsGKjIMyKqNJHFWLhGsGzu4koXtXi/dU9cu/7AY6tpnQAT
                jb9C9oUMz+nPbvVFoJr5wrg8v+HXfYVidEsAw+MAVdEDIGmZ/IYr8oUdsBcgoorJu4jJhICgjBxJ3wFI
                v6+DVR9NdCUeL0e5aSv9HriHx2lVPbnPR+MPAGz8aQpYhIVrrg2ioGrqCQi88U/rKhUZfvPP82YvgpFh
                gXFz/r3yAwJWFQEArBxd3IralJW6ooirVa3GlqbTAOwz6YMKBgD9nf2vQ52zAHit8ve/dVbdKWUseh/X
                T0Wf6evu++9K3TdRSAUaBDiWHocQBwFBNv7Rb7wSc9K6ysFw4z9uKl9WIJAtKxAY+9rYH7gdsA+Ok2eM
                W60vzG5p2VHmdkcEmntakOPRWJYo2Z1caVlyCIC1OUuHfH/Aqjuss7MzWbZCFTG3jwWy1u+liKaYRbho
                zXVBFLRk32cfD2sQEHTjrzVYBbH2zflmP/qHTGKfemX+50oKNKAqAgCRfEluWlcrqctN1y8Idtz+HID3
                5HwKFewBGNG7qfevya7EwWnHOgDQ86FyhQhOwZC2Jbt6L9i6aVNfOcvTsWmfOTkKf8MSRFOR4EYDQUBo
                9tII/M2/FquQPcV5Yhf/uAZ9LBDImfkPcCngAuXNclfFFxtaG9oKuVgVskTFbeqjPbNtZvmmy7nY3NPz
                t2RX8gfJ7t7bEp2JnyWTya6KFKSW674CovJmEPdLFFrBBwHHIgxBgOKaoBr/hm91tDm18gQg78v3Zj9p
                bf9cmf8Tz5kYCBhSDQFABBC3bv5pko5M6VyAaCz6WZ24s9QEtUM7VbwXIFCq7hsLicMAgGhHCwIU1yw+
                YF0gOyY2fKujTZzUaoz2QA9n8QtyzPnPn/DnNiwwNh3AjNAHAE1NTe2Auq52J4LzGtsaW03XtUJEAe8l
                QS2taB5A4Cz36UcissV0FYlCQXAjFj51fRBFGQ0Cgmz8b9+0mzqp1QrsNb6Lf/gP4vPN3s+wgEGhDwAQ
                8bXIzXSkrSmZC9DY0nQSBO/3Ok4CyAMIlE7adnicNGSz6SoShYfeMKWDgIAbf6lJrRaRvUbe0MeN9WPk
                Dy5v9sifH5B7QSAzQh8ApFV9rXInwAVTsBdABOpvf3BHp1QAoB4/HjUhzEomMmuKBgFBNv53btpNa1Or
                HWDP0S/KWGvtlvA3ebe/rHPGnecybTBgoQ8A8s8AmGS6pK3LTNe3nBrjzScCOMjXwVLZqYBBE6h7D0Ba
                0qbrSBQ+egMufurWIEoKJAgIuvFPp1cDsicEOcb03d/sgVz5AeKZH4CKbefmLfQBAPwHAADwj/Zudtx0
                lct26+oUsh3o3ghFp1KZmF4hg6haqV4ZZBAgap0Ewbby30ewjb/jpFdDRt78hxP08iT8ZaqXe1ig8PwA
                c8IfAGhBAcB0pKxLTVe5HOyYfRyADxdwygzbtmMFHB9uXtscS8oxXUWi0AowCFh0wDMPi2N9sqxBQICN
                f/3yzt0dx1kNYM/J3fJjDXquN/tMVSeO6ReYH2BQ2AMACyPb3vqmX7Ztu8V0xUsmUvCWx1IrUygPQKa5
                fipiegotUbhVaxAQcOOvkZE3/7HWOFeDnjcQKEd+gCGhDgBisdjuAKYVeNp01OES03UvRVNL08cBHFro
                eSpTaSqgFvrvTkQTqV6JhU9+PYiiyhIEBNn4f69zd6cmvVqBPcYa7/ENet5AAOXLDxgM4mbzCPV2wEOW
                7lXUe57Kl5qbm2/r6enpNX0PxXDyrfnvfeKU6AFob2+f/tbA23u6HaNqOn+2eA0NDTOlTt4rIntDrL0F
                GlVFHQSzAEAdbBWRFATd4ugrYsmL263tL2zp2FLVyx83tDa0WY51BIDDBLKnis6EyiwAbwvwmgO8JqJ/
                Tm9LP7h582Yu9FQ2cgUWPgksO/TKSpe06IBnHr7uuYM/qZbzGyh2KujkoBv/lLMakD0gCmT+l7V970hr
                raMN+ti2v2N/UZ3YzgugOtrQi2B8X79kPhvfrpnLAgx1ACA+pwDmsEta0pcCqPg3fLk1xpuPgjqHF3Ou
                SGWHAGa0tjZOx+Bu6mgEAo2kI29EIpHuzs7Od4q4nNUYb9xHRJwBTOvbOZWaBQBOxIm/Nfj2hfDchzxy
                MIBq2Q2wNhqLHqUiRwN6BIAPAMPrHGT9shj5g4z8ptHhtwZV1KWnqd1iPy8if3RUHpkxbecHN27c6LVD
                YxjU2nH7s1BrIRz9yLhbzYrhFDgoc9tyUWR67WC0JfoYFHcluhO/BMAZHyULeRAQYOM/61+6351Op1YL
                MLzC7PgGffgrOQOBsQZ95Dwd+TgrEPAIHgDocARgeiAz1G9Rdrzpm1AtNqnv7Yhj7VFtvQB2S/SPAD5a
                3Nm6PtmV3L+4c91FY9GT1cK9UNRO+CilwHV9XYmC1uaOxqK3qpQUoDmAvg5IH3wMpSnwh76uRHE9K0Wy
                4/aB4sg5KvgcgGhZLy54U4FfW4q7E12Jx4O8L7+isehnVbAMQPHrcwheAvTWZGfyblRjILBwzV0AzjVd
                jVGC27B0biBLp1/33MFH+wkCVPG1mw5YtySIOs3+Tle7o84qQEdzyyY3wjoWjGd9KfvXjIz/69hfJi5g
                knWeTDh05C+q1v79l7xrfRD3P1G4A4AW+9eAfCrPx9sA1MKlF0MUX090J64yfR++7zdmHwGRx0q4xPZk
                V2IXAGXPkG9sif5CgM/k/lQGd5228yy/b6QzWlsbpzmDm1B4fkdparQ1+Xqys9LFROPRwxS4EorjEMzP
                2DpVWd7X3XsvQtBIxmKxaMpK/ysUC8p1TRU8B1iX9HX2PGL6/goStgAACFUQEGTjP+P7HftYaVmlQOuk
                bjdMDAS8G3QgTyCQaxUzl0BAVYwFAKFOAgRcx7RfUMF/uJ2sgq/EYrHyvnlVkmWVuqnH9OHEyfJXDZiR
                /1Ot27Ztm+/1F6Y7g/si6MYfALZVNum2Idawrx2PrlLFE1Acj+AC7ANF9Md2S/SpxljjR0q/XPHq4/W7
                p5BeU87GHwBEcYCo87Adj943a/dZ9SbvseopLsfFa74RRFGLDnjmYVGcBCDXy8GVgTX+3970HqRltQO0
                Tl7bf/ixjMvW9074A7ISBZF1gEjeDYNGzxlftDFhDgAEkD1cPn2pJm0thmDI5Rq7pCV9sekb8SMajx4G
                1fmuD0RwCoA/uR0zaDlGEgFTkZTvTBYHTtJEHVGh5YPb2tp2suP2Nywr8hco5hm6NwD4sIj1lN0S/UFD
                Q8PMoAuPtkX3imjtEwUu3lUYxcm1Q7XPNjY3/7+g729KUVyOhWtuC6KoRfuv+4NaegKAkQ28HKhcvnj/
                dYEEIbve2fk+WLIaIvGRRnd8tv74hflzNeiegcCktf0LmDZoUGgDgIbWhlYg/9iRKF7q6el5GSp3u11H
                ga/MaG1tNH0/nhTXehyxNtGZuA8qv3U7SAzNBIhIxPcbfV9X3wsA1gZdR8uydi73Ne1We5/t6YGnoHJ5
                jvwIEywAX7SmWf/T0No0J6hCGxoaZmoaDwFakR6o8eRdYjmPNcainw/q/qaoy4IKAm7a99lHHEvfD5Gr
                HZEjFh+w9ptBlLvLdzs+gBr8UYHhtWHGGnQgfyAweREfGe0UmLyan49pg5I7eOA6APkqlrZc3yJU9CUA
                SMvQEkDcplLOqHMGQ70uQENz8yEKfMz1INFFABRwnnc9rJJvXy7S6XRhc1lETwewKcg6WpZV1tyIaCy6
                AI6sRSarP2TkXZaDxxpjTV8KojRrWs234X/RrreRmcHxewh+DugjAF4usMjpIvhRYzxadTN9QiawIGDJ
                vs++uni/tbcu2W/tn0q/mredv7PpQyKRRwBEJ7+NTw4ERr7sZ1hg0jkFrB+ASeeZE95pgF7b2zrYAACb
                Oze/Zrc03Q3ggvyXwoXxeHxpZ2enqa5nV5bleI39/yXZmXwAAFT0ZXHrN1LsCTMK+k5Odib/PrNt5vvr
                UnXnqCX7C7DLcP3fBOT9gBYxli2DgI5sTDI7q06OAPd3d3eX7d8/Go9+VRXLUFwQnQLwNET+DHX+DsUG
                WHgzko68AQBpK12vjs4SS/YGrPcAOBTQA1Hwz6vWieDbjS3ReF9XwquHqfhnEYsuUOjpPurziKjckehO
                PAJMHrpriDfsJk7NPLH0bL9DKQIsbmtr+1ZHR0f516LfcVyGhWuAZXOnzJbqO/+g8yBReQhAw1gLrtmz
                +oZ/OWSm8mUl/g+3yVkHYiwIkBzrAAx/Jf/6AV7TBg0KbQAgKnu6LfXiWM5LI39Oy9DNEa09B9B8b6G7
                DiJ1IYBSk+zKzo7bB0LxCbdjVORGDH9bzZg245W3Bt52kK/hMdQDAK0p+Pt5eGGb2yd+vb29ffpbg2+/
                DEXeJZ1F8GeF/lONU/P6cMMeWG+aHW+6QVUL3XY1BcWDlsiPNaV/SCQSfvIRVo38oaGhYaY1zToBkNMB
                HI0CfnYF+Jrd0hRNdvV+qRLPycd0zrcgen6yM/mfbgf1d/a/DuAeAPdEW6KHK3ATvKfE3s/Gvywuw0VP
                CZbPqfodVXe+q/NgSctDAOrHt9dZk/gz6/VkBQHA6Nobw9fJFwiMb9Azn+UOBLLOGb3E+AWBTAcCoR0C
                8OgBGOjv7B/tPt7cufk1AD9yvZ7qRTPbZjaYvq1JHLkBLt8DAvytr7N3dNx/eKpdl8uN7gkD/67l/D4e
                vkfXHIE05LpkZ3Jdd3d3AkE2/rGmy1BY479NgeUacdqT3YkTe7t6f+mz8R+nv79/S7IreW+yK3FsxLH2
                AfS7yJ1ZnYdeYMfssq8Lb7fYBwE40OWQtxxLPubV+E+U6Eo8kexKHAmVK5BnaqMAv0p2Jnz0PJAvopfi
                oqcCGZuvlJ3+redQSVuPQqQ++zfSWLd8Vt+8j2GBsTa7mPwAf8MCJoU2AFD3XQBfwYS57j5yAWbWOdMv
                Mn1f2aKt0Q9CcLzbMQ7kOkye1/+SyynT6uP1babvrWQqrkvBWnBeC7pKdtw+DaL+M5cF96FG9+rrSizs
                6+grW75DT0/PK8mu5JcijrUvFA/4r49cYbfYC8v5TETlbLfPFfLl/k29TxV5eU12996mkAXA+PXlFfhD
                oitxCuA6C4gKVcVBwPQfdB4O6IMqmDk6417GN7K5GnRf+QEFTBscPWf0yx6BgLmVgEMbAMjYvsw5P53U
                AA73AvzY/bJ6UZjmEKuj18M1BtT1fV29v5n0VXFPmPJKoKyImnKvze+4vtWntDbQZXAbY40fhlp3w1/M
                noSjJyQ7E5+r5MJDPT09ryS7EyeKYgEEPtfOl9sbmxvn+zvWmwrcZho80dfV+++lltHX1fsriB4G4Elk
                Vp78hZXCZ8HGvzJEL8XCJ28v/ULBmf7DxBGWRH4nwIzst/HhGxqXbDe+EZ4cCGDceVnnjB4+edrguEDA
                z7RBAGHYzSSUAYC9m92CkaSwHFRzvwGnZegm114AxazawdoLTd8fADTEYvsBcqLbMRas65FjVT9x1DUA
                EEsCDwBq1HRnVuXYtj1DYP3EJccki/xXWlIHJXuSK4KqX6I78fNI2joQwF99HG6JWP9elgWy9kMdgLxL
                TytkWbnuMdmZXJfsShyW7ErU9nUlTi5mGIUKIZdUSxCwyw97Pi7qPKiCXV3f7MUrEBj+ut9hgWLXD8h7
                TvBCGQBIyr0BsyR3AJDpBVD3Nw5LLp7dPnu28XuU9A1wff66vrer99e5T/aYMmVmJsCUDQC0Tr7lJ7lS
                gIcjjswf7o0KVE9Pz8vp7akjAHjvCyCIp5C+q9Qy65PN+yD/io5qpfShoJ8DlVP4g4Cd7+o9xlHcj+E1
                Y/K+2WNiIDB2jZwNesHTBv3lByBXWQaFMgDw+mWrjuYdA0+hZrFXL0DN9pqvmry9hljDvgJ82vURqNyI
                PGv6O5bl3gMgwfcATFV2zD5CFGf7OPTBWbvMOqGnp+dtH8dWxObNm9+cHpn2CQCPeR4sOMGO2yd6HufC
                cumlA9DDt/SpQC7BRWvuMF2LXKb9W+I4R/AriEzP+Wafqf/Yf8YFAv4S/vIGAhPP8ZkfkHtYwJxQBgCO
                xxusWvkDgDe6ul5VVdc9AiByicleAAuRa+H67OX54W1Qc6pNiVsSoFcCZUWo8c6siohAZLmP49ZJCgs2
                bNgwYLrCHR0d25yB9CfhZzhAZWl7e/v0oh9ORCP5r13+DanIEMHFYQsCdvph72clor8GMD3vm/1oIzy+
                QS9LfkA5hwUMCmUA4LGanTOjbsZGt/PTUrPIY0bA7NqB2q+YuLdoW3RvCE52v39dBJcd/Ya3ON6a/wLY
                C0H3L03BHIBoPHo2vFf5S6BGTwjT225/f/+WiGN9GmNrr+ezx1sDb3252HJSbiuZCmJtbW3+9oKn8AtR
                EFD3o94FKvITALV53+yz5AsEgsoPyD0sIBMCETNCGQBA3dYAkA6vbWeHewHudS9CLjaxYYqm9ToAEZdD
                Xkh0Jn7u4xm94vLpzsOJlFQ8UcWl3gfhy0FsMVyonp6el0Vxno/bvGQ4ma9gdWlrg8vH1vbU9rLNNqAQ
                GA0CzPX21d2dOEUg96qgdtKbPbIb4dyBQNH5AYVOG8xxXt78AIPCGQDAbQhAfa0ZHnFkCTJLrua7TkNk
                WiTQXoBoW3RvQE51PUh0MeDdfSriuhYAMBT4pkBTqgfAjtvHA3ifxy3/JNGd+IXpuuaT6E7cB+B3Hoe1
                RjdHzyjm+sOLMCXyPx75J9PPgMpMcDEu/vMtJoqu+1HidMlsAV8D5Hizz2qj8wUCRecHuAwLlJ4fYE7o
                AoCmpqZmAG5v5hv8XKe3t/clzXyzuLnEtu0Zfq5XFg6ugfvb/4vJzuR9fi6lIl5TAQOdCaA6tXIARMUr
                UXRbWoauNl1PHzdyicdwGFRR9IZBCqxz+fiYppamz5h+BFRmqpfjkqf3CbLI6XdvbofIXRCJ5OziB5Av
                EMh8FPb8ADNCFwCgxmMGgNebb5ZISm6CSy+AAo1aF8xbSnNz8x6q8Fi2VG9EnmVPJx3puE8F9EqkrIAQ
                fDuXR1NTU7MC7t3XoneamO5XqGRn8u+A3uNx2MGZ3qnCiahrD4gD/VE0Hj3M9HOgskrBGdha+mUKZo17
                s8/+lSMTAwHkf7PPkjc/AO7DAiXlB0wYFnCNziv9QA2WnVPaUfc1ABz/AUBvb+9LInDNBbAUlwXRC5C2
                nH+G2wYuig3JruTP/F7PEo/FgExtCjQFpGv0ZLj31KTSSP+L6Xr6JvpNeAwraUpPKebSqbrULwC4zX7Y
                VRV/iMaix5h+DFQ2V2LZR7tKv4x/28+p36iWXAAZ25pnbG3/LPne7IEC8gMAr2EBAMXnB+QIHkwJXQDg
                1XWt4i8HIOsOXXMBhnsBis6E9iMWi7VDcJbrfUEWwzVnYYKIRyCkzAEo+kZUP+lxo7+thrf/EcO9AKtc
                D7JwXDHXfmPjG28A6jVstbOK3N/Y0nSm6WdBpZIbsGzuMhMlD53ZeLdCzoOMBbNuDbpbfgBGz8s6B+7D
                AqXnB2SdM3q4wORmAKELAOAxhz01LeW7BwAAEh2JFwF13YlMFJdFo9FdK3VLKXGugqI2fwXwUl93708K
                uq9ZiVfhPlzAHIDi1AByiPvN4qemK1kwgXvvkspBxf4MaESvBuCxAJLWCfQeu6Xpe4DLzwKFl+J6LJtz
                o8kqDJ3VeBdgnQlB2u+bfbmGBQrLDxg7zHNYwKDwBQDIPwQgQF/mjaMwEpFFcH+7tjUiRSdCuWmIN+wG
                6Nmud+ygsLd/AFiPQUA7XI6Y2dzc3FSJe8qppmZKBAB2i/0BAG4N4Xak8aDpehaqDnX3w30YoMaptQ7x
                eblx+jr6NkH97pKo59tx+4nGeON7TT8TKsh1WD53kelKAMDgWQ0/AawzAEnlzfzPJj7e7Ms9bbCAZYVN
                CmEAkL8HQIHCuv+HJToSG7x6ASB6WXNz8y4+L+lbBJGrkH+99OG3/8S9/q+Yfa7lviJgRIMeBqh6qtZH
                PA5ZG6ZFf/zq7OxMCvC8x90fWOz1k43JWwH80dfBKv9P1FrX2BK9upSVCCkgqtdi2dzFpquRbfCshp9C
                kQkCgBxd+QW82SN/IJDzzd7PsAAAf8MCwu2AR8TjcRtAfd4D1H8C4EQS8Rxjb0pL+h/LeT+2bbeo4hzX
                egEe6xW4UPd8iLQT4DDAFFkJUCyvZyZPm65jsRTypOudlTJzZD0GB6y6zwB40ecZOwlw81sDb6/nVMEw
                k69h+aE3ma5FLoOfb/wZBKcrJDVxnB/wOyxQnvyA0S8Ute2wOaEKALan02WbAjhRJhdA3MdBRa4qZy6A
                1spVGN6lKs8Rrybqi3z7ByBwXQ3QR2NWPlNlLwBVbXf7XFSfM13HYgn0Bfebd0r6ftm6aVOfRHCsAhsL
                OG0PB/oLuyX6eGOs8cOmnxFlEbkGy+YsMV0NN4NnNt4nFk5VYEgLWBCo7NMGZWLw4J0fML43wIxQBQBi
                aVHbAPu+fkRvhGcugF5QjnuJRqMxgftSrCKyODOWX+wNec4ECHIq4JQIAATY3e1zJ2K9brqOxXLE2uh+
                89JaahmJjsQGGdK5AP5U4KmHi1hP2y32PQ2tDW3mnhIN+2csnXOz6Ur4MXhG4y9EcCoEQ25v9mHNDzAp
                XAGA19z1EoYAgNFeAPcpSyKXx+PxnUu9F62Ry+H69i+vJeoT/15KGY7jeOVEcC2Agqnrv72mhwKd/1zW
                O0ujx/0IKfn7HgCSyWRXsisxT4Br4b5GwKQKAHKm5UResGP2pXBbN4MqRQFcjGVzjSz3W6zBMxt/6UBO
                guhAzjf7rJsre37AxLKKyQ8wJFQBgFcXZNpKF5UEmE0iegPcp881D+pgSbkAmVwGPd/9KOemkt7+AQxG
                pns8j+CSAGumSA4AIK4pOZGaiMmFu0oSiTgejbGWc/e+VKIrcZOjkYMgWmjexK4Q+aYdi65tijfNDfo5
                7cAUgotNzfMvVfrMhpWOWCcB2D7WLT+5NyBzoyhffkDJ0wbNCVUAICJub6zb+zf1l7zrmq9eAJTWCzCI
                wSvgOpVMXks2JH9c6r1s3bSpD8AbLofUz2htbSy1HD9UNVTfS0Xfh8ccdWvIqtp97tOOZ/AyzdeFCtDf
                3b0+2Zk8FNBLAbxT0MmC9zuqTzS2RJcVu2Mh+aaAXISlc5ebrkgp0mc0PKhAJggAst7GvYYFzOUHmBSq
                X9rqPmb9CnzskuerHKS9dtyLDepgUbkA8XjchnquKXBLqW//WVw3R6pLpTgMUADx+B5zap1Q/cwUwrJS
                Xg28r30oipBOdiXv0IizDyDfL7AcS4CL7P7ok/Xx+t0LOI/8UyguxLI51bO8tYvUmY2/V8WnAGwDMOFt
                PE8gMDo2n6WIYQEAReUHmBKaX2az22fPBmDnPaDEBMBsfV19L0Dh1QtwRTG9AEM6dAncF5J5ffYuM+8u
                17145kVYXAugQK7d5OlUumrfRMV7jH+okuX3dfRtSnb1XmBZcpAADxV4+sERrXm6vrl5/8o+pR2OAvgq
                ls/9V9MVKafUWY1/UMGnANk2+sV8b/ZZD8JXfoCPWQYF5QcYFJoAILI94tpQaYkJgJOuJ84iePcCnO/3
                egAws21mgwJfcb8PuXXDhg2FJEZ53Id7D4BI4HsCVLvtbh+K1ga3fXSZaVpnexxStu9LN72bev+a6Ep8
                3BH5mCieKuDUWMRyHrZb7UC3op3CFKJfwbK5dwZS2gPzT8CKIx/CinkV3XtlROqMxofUkk9AZPzCXS7D
                ApmHgtLyAyae45kfYE5oAgCxxGsKYMkJgNn6uvpegOAX7kfJlYX0AtQ50xcCyN9ACLp2qqkr39s/fEyN
                DH5ToPLdmQECJNxr5VRtN7RE5N3uB2jJOTaF6O/sfTjRnZgrimMB/N3naTE48oht2y1B1nUKyjT+Sw/9
                diClrTjyFIjeD8jRAO7Einnfww03VLz9SZ1e/7g61idV5J2J6wB4DgsALsMCQN43+wJmGZgOAkITAMBx
                n7Kmjpa1BwAAHCd9Izx6AYZ06Dw/16qvr58F1a+63gNwc0dHxzY/1/NN3XsAAK3WtyUjAYBCXBdXcizs
                YeZxlOHe1KvuYmSHw0R34sFkQ+L9UL0O4msYYjfUyl0IxShqVVJA/imwxv+B+ScA8mOMb2/Ox8F//E4Q
                K4imzpq9ChY+rsBWLSDhb/hBlW3a4Mh5k5YVNig0AYDnNsBW+QOA/u7+5yH4pWu5wFVtbW2e06Mi0yIL
                AczOf4Po2smadle570GH1CMAAAOAAqhXT5OimlerO9j9YzMBAABgPQaT3cnF6jiH+cz3OcZuaSrLol07
                GIXIl7FszncCKe1384+B6M+Re8X787Fy3g8C6Qk4reFPsHAsgK1uCX+ao0HPPDRUdllhQ0ITAHjMAHBm
                1M3YWIlyHSfi2Qsw4Ay49gI0NDTMhMhFbscIcGvZ3/6RWXQFgNvmNDOampqay11uxRnaXlgcZ73r54Lq
                nJeemUbnHgCo86zpavZ19/3XUM3ghyFY7ePwJZXcxnsKciByDpbO+W4gpf1u/jFw9Ndwn176BRz0eDDD
                Aac1/AnAMQC25nuz9z1tsKzLCpsTmgAA7qvWbdq4ceN231cqQH9393oFfuV2jFcvQGRa5EK4bWIEdE+z
                pv2gEvUfrqDrW6tGqnYYIHCDNYNPwW1oTtESbY1+0HQ9C9WYbDwMgOvOexEnUkhCXsW8+dqbm5P1iU8o
                3HvnAG1ALc41Xd8qkYbIuVg6p+T1R3zx1/hniH4RBz3+vUCGA85oeBIR/QSALUCeLv6S8wNyJ/zl60Uw
                KRQBgG3bMwC4vaWWNQFwItXIDXDrBVC0DDgDX8z1UXNz8y4KXOhRwtcr8fY/QjxmAqA6ZwIY+SnZ0rGl
                H1D3bXMdLDDyREpgRSyvHfd6e3p6KvpzVpD1GJwxbZczvFYR9P7ZIwBpKMLZ+I8Q/SJWzgsmCDi1cQ0i
                egwkEwQAWYGAn/yAieeUKT/AhFAEAKhDoFMAJxruBfi1ax2Aq3P1AqQizlcBRF1O7amTuu9Xsv6qHnkA
                1TgTwDI3OKaQRz0+Px3VtE79fqhThWsAIIIHTVdzoo0bN253nJovuCYGKvaMtkWr7/s7OGlAz8HyufcE
                Uloxjf+Y84IMAgTWfAg2Z3/dV35AuacNGlxZJBQBgDieUwArGgAAgGrkehTYC9Dc3LyLKC72uPI3Ojs7
                C1sCtVDez6cahwCMBQAR8dg2Grq7Hbc/a6p+hWrsi54K9x42qKM/N13PXPq7u9fDIzh3UnKU6XqGVBqq
                Z2PZoSVtOuZbaY3/iPOwct53gwgChk6fvVZgHQ1Bv583+8rlB5gTigDAEfd961W04l2T/d3d6wW437Ue
                iivb29tHx1HTVvo8AE0up/TWSV3FE27Uce8BcKpzCMCY3s7epzwz4lWuREh+fjxYIl5BKrbM3nX2I6Yr
                mp/+h9unIvig6RqGUBqQz2P5of9R+qV8KE/jP+J8rJwXyBTBodNnrxXHOhqQ/lxv9rmGBcqaHwCUbU34
                YoTjF5gGvwZALinHcu8FAFq3Dr490gsQAcR9/FHltoq//QNIW3Wuz0cyzzcc/9Y+qdndBRUCr/HSD9px
                O/QJaI0tTacD+IDrzSp+XM7VKcttQKatcftcoFG/19pBpCE4C8vm3BtIaeVt/EdcgN/NXxpIEHDm7HXi
                yNEK6c/ZxS8VzA/gLABAxD0A8N72tjw29/T8TYDfuNZVcVV7e/t0O2Z/DYDbymqJiEogc23f6Op6He5L
                2O5UH69vC6IuZWT0R8NK4U54LAsMlSWxWCy0jY9t2zMEerPHYemII8tN19XN8K6XW/N9ru69cDuaNIAz
                sXTuTwIp7YF5x1ag8c9QvQgrj7wjiNsYOnP2OsvCUQr0aZ7GuXL5AeaSAEIRAHj0ALwx/AsgGBYWwX1o
                pvWt7W/fB5HrXW9JcHtPT8/bAdXaQWa3xPy3JXVVNQxgGQ4Aent7ewDxGjttSiEd2hXptEaWA3AN/AT4
                bW9vbyA9bCXKv4OgSGh7LwI2BMgCLJv7n4GU9sC8YyH4FSrR+I+ShVhx5NIgbmfo1PpnLQtHA+gbm7sP
                X2/2pWw7bJLxACAej+8Mgdu63oH+ckpsSvwF0N+6HiQ4AS6/9AXok0ENZpnNER5LAovDXQELFXHkFnj1
                AghOsGNNl5qu60R23D5VBOd4HJZKa/prpuvqZXixn1l5D3A04f9qU9YggAVYNudXJV/Jj0Aa/xGyECvm
                BdMTcGr9sxZwFCQTBLi92ReVH5CrF8Eg4wFAykrtBbc3qABmAEwqUuRGlJCg6QC3J5PJrcWeX1ylPZ9T
                tQUAxn9Kenp6XlHgdu+a6jcaW5rOMF3fEY2xxg9D5d88qy34Xn93//N+rmlUDQ6Ey/eDAr2mq2jYIIDP
                Ydnc+wMpLdDGf9TFgQUBp9f/xVIcBUHSc+fA0Qa99PwAE4wHAEi7N0yilV0EKJdEZ+JZqDxQ3NnSL0Ma
                +N7a6hUASHWtBqjGO8cyahzrFgAdHoeJQP+tsaXpJNP1jcajHxKxVgJw3cVSgL5IZhns8FOc7PaxBawz
                XUWDdoTGf8TFWDnPOyAvg6HT6/9iWTgKyAQBOQOBYXnf7L2CB8B4EGA8AEiLus8AkGBmAEwklt6AInoB
                BHpH4G//ACyvXQGrbzGgUAQAPT09bzsi58J9dggATBPoz+24/UU/162EaEv0cAVWw31hKgCAAzm/u7u7
                5K5zu9Xex25perqxJfpkQ2tD2RNNbdtuUcHZrgfV4E/lLrdKDEJ0wQ7S+GcoLsGK+d8MoqihU+r/mrbS
                HwWkG8CERXx85gf4GBYwyXgAIOq+CyCc4IcAgJFeAKwo8LQ3hqYN3Wmivoh4Lge8B4BaI3UrhlqhCACA
                zL71EPXz5hGByg/sFvueeDy+s4/jy8Zusc9XyCNQl7HyMXf1dfWWPFY8a/dZ9XDkN4B+RIC5lhN5qrGl
                8X1lvC3RWvkWALcNfzoSHQmvHTGnokzjv/TQ35R+KR/C0PiP0kuDCgJwiv1C2orMHxcE+HizL3TbYVOM
                BwDw6AFISa259ckL7QVQveONjW+8YaKqiY7ERo8lU2ujbdF3majbVJDsTF4D4El/R8uZQzr0dDQePazS
                9Zrd0vIuOx59AJDvAepnPtG6iGNd5OM4V3vttde02qG6XwN4b9aX2wTW43aLfVAZbk3slqbbBXBfcVG1
                ostsh9QgVE7eMRv/EXopVhx5WyBFnTLjhTSseQC6Rr/mMSwA+Jw2aJj5AMB1CqAMvtHV5TX+WjHJzuQ6
                CFb6OljwZmp66l9M1RVACg5e9Tii2oYBwmRoMDJwIoAX/ByswP6qeNxuse+pxFr1s3afVW/Hm26oQep5
                KI73edrrqNETyjA9NbL57TfvBXBEjs9sQJ6wW+xLUOTvl7a2tp3sFvvHgHqtYPhOnVUXzL724ZFp/JfP
                +W3pl/IhlI3/CLkMD8z7RiBFnTbzf9OIzFdBV77Mfy162qA5RgOA4WV1W/MfoRvhNv83EHo9/PQCOLrU
                1Nv/aE09EwGraE8AS80HpxNs6djSn0LNMVB0+jxFADlT03jBjkd/1hBv+hiASCl1iLZGP2i3NN1RO1T3
                KlSvh0eyX1ZF+iKOdVzy9aTfuuevQzx6tsB1c6GdALk9Go8+GY1Fj0UBrzpNLU0f354e+CsgZ3rflH67
                s7MzWer9VBE2/hMJLg8yCHAikXkq0qk5GnTAx7BAjvwAk4zuaLZl+/Y9LXENQowvUJLsTK6zW6IPAjg2
                70GCN4dqh75luq6WYIMqPu5Sz2rqATDfP5bDG11drzY1NX3UiehDAPbweVoEigUWdIEdj3apgwch8lgk
                jTW9vb0bkT/ItWa3tOwWkfSHoHqkAB9TB+8rOHVI0ZlG5GOJnu715XgGjiNDIt51UMUhEKy0W5qehzo/
                dSLWI/2bev8bQCr7uObm5j1SEedYUXzegR7s85421End9b6OnRoGIfJZLJtT5OykAlVD4z9CcDlWzlMc
                t/rKipe1YObfnfu2zLfSWKWCOBSQkZ/HkSAgk96f+evIz0n2f0SGv2w6BdBwACCW7uX2DKTC2wD7p9cB
                cgzyNkqy/M3X3txc0CUrUUvVl1zbzeqbCRBKvb29L0Wj0UO1Br+Hxzr7kyhaRHAuoOc6EcBuaRoE9GVA
                tgDO8P7k1kzA2QWQPYHU9JJ+Tyg2RNT6eLKnu2y5NDvV1P18e3rgQgA+x/p1X4gsshxdZLdEt0PRAcGb
                yPSG7JGGM9NHPJEtLYIvBLHPRkhkGv+lbPzzUlyBFfMVx6+6quJlLZj5d+enW+ZZqqsgaFVMaNDzBQJZ
                3+Mqw5+o2SDAaDer1+p0QewC6EeyK7kWwO/zfLxl0NoeirXULViui7o41bQtsOnBMQ+JRKIbQ3o4ICWu
                ua51AN4L6EcAOSrzf/0IIPsBmF7SpQUrhuoGP9LT01PWn6OOjo5tksLxKniuiNOnI7P3x0EAPghgZqEP
                DNAvJboSj5fznkKMjb9veiVWHhlMr9ApM//PcWQegE1A/g2DAB/5Aea2AjCcBCiO+xRAA6sA5qfXIkef
                jQLf2tKxpd907QDAqXFcfyELsHv2dsbVLILBQKfZ5ZJMJrcmu3pPB/QCANtM12eMDKrgqmRn4sRK9Uwl
                EoluZ1vqcAhWB3lnClyT7Er+IMgyDRqA4jOBNf4r5h1XvY3/MJUbsGLedYGUdfqsF4eDgNFE9bFAYNi4
                Of+58wNMMpxoJe91/dh/slXFDfcC/GHCl7cORQYC2ajCVx0zCV5uSVHW1oGt7/Z7vWw6Ydx2Ikedcm/I
                4v4j4rjuHxGoZFfy+1ZaDvA9Y6SyHnc09aG+zsTXUeFBxs2bN7+Z7EwcjczS2ZVO1h2A6Hl9XYlbKlxO
                WAxA8Vksn1voWiTFWTHvOAC/RDU3/mNuDDQIgB4OwcbsL6vI+E5MH+sHmGAsAGhra9sJwCFux6SdyHaf
                lwuGjl/iV4EfhuXtf4QAv3P7PCKRhiJv3vU+HdSUez12165hK2KFahve3t7el5KdieNVrE8BYmJ9/RcB
                PSPZlTgy4PX908nO3hscxzoMwDMVKUHwkqpzeLIz6bm/wRQxAOhn2PiX5EasmHdtICWdWr/RSeuRkAk7
                suYbFpgYCBhkLADYvn37bLi9oSg6Z+20U6hW+Ep2Jx8C8CKQmVaVRk1o3v5H6JBeBeAx5H4j63YiziuF
                XXH4urDcVo57cXNPj6/58f6J66qF6mh3ecsrj77Ont8ku3oPUJGTAPmvSpengucUcmayK/G+ZFfyXhhK
                Le7v6flzsivx/yB6GoC/lOmyb0D1smR9Yt++7r7/NnFfBmQa/2WHBtObNDUb/xGLsHJ+MLtdnl7/qpPW
                eRB5JdeCQG75ASYZH4mIxWLRAdVmC6kmRGosOKktYklKhuT/EonEW6brl6u+KSt1iFq6rq+jb5Pp+uTT
                3t4+fdu2bTMxvIyqiDjd3d0d8OjKd9MUbzpUHY1DsFVVt1tiNQJAjdT8sdzzsaMt0Y86is+LYB8BZmtm
                GeOkCDYAeDzRmbgb5n9+vJ9ZW9P7nZRzOkROBbBbmS7bq4qfieg9ya5kZd66SxRtiX5URU+BykkAYoWc
                K4qnVOSe1LTBn5peW6NkC9fcBeBcn0dvg8iJWDrnkUDqNrUb/zEi1+K4VTcFUtZ9/buLE1ktwB6ZfNUc
                1ZmQ+R+Bs//gaY1lmaJb8KMxUSjRDkjqm5v3q4k486A4UgUfgOJd8J6Km4ZgIxw8D8jjEtFHEpsS/wPv
                zYlCw47b71FHDoFgb4G0Q3Q2VDONjlrvwNLtqngRkLUqqWf6O/tfN13nsvEfALDxryj9Go5/bEkgRQ0H
                AQD2EIxfB2BE9rRBBgBEO6Zau9V+N9JoBbCzQHZWS2sBbBWVN1OO1d84Y8aLGzZsKHeCJQXFXwDwDkQ+
                yca/4q7B8atvDqSkTBCwCsCewPBiQTl7A4AI0sYCAKMLARHt4IaSm5L/B+D/8h2wuafHdB2pst6BgxPx
                rTmPBlLajtv4A8ASrJiHQIKABQ2v6X/2zROrZjUUeypkeBXgyQsCDRocyAzdeutERDuI4cZ/Lhv/4CzB
                A/OuDqSkUxtfVyc1D5m8pbwLApnEAICIKHjvAHoCG38DBDdj5fzKLxkMZIKAmuEgIGtBoLAEAgwAiIiC
                lWn8lx26KpDS2PhPpnpLYEHAZxo7VCKHA3h+4sqAplc8ZwBARBQcNv5hoXoLVhxZ+R0EAWDBrt2ajswH
                kFmka1y7b24zAAYARETBeAeWc3yAjf+nAfwabPxdyK14YN4VgRR12q49mo7MV5FMxj+HAIiIdgiZxv+O
                w4LZPCnT+P8UmQW0yI3gVqycd2EgZZ22aw9S1j+MBgGGMQAgIqqstyDOJ9j4h5ZAsQwr5n81kNJO27UH
                NTJfRf7GHAAioilLkxDnWCw97IlAimPjXywBdDkemP+VQEr79Ixe1Mg/QPCcwRQA0yMQRERT2Mn3RfDz
                BZXeKjmDjX85KFQuxAmr/rX0S/nwq61NSEGxYEbCxM0yACAiqnZs/MtJIfpVHPfYnaYrUmkMAIiIqhkb
                /0pQAF/B8au/bboilcQAgIioWrHxr6QpHwQwACAiqkZs/IMwpYMABgBERNWGjX+QFJB/wvGrvmO6IuXG
                AICIqJo8cOSnIHIf2PgHyQFwHo5f/UPTFSknBgBERNVixZFHAfIAgOmmq7IDUqh8GSes+q7pipQLAwAi
                omrwwD/MgTgPAdjVdFV2YArRL+G4x75nuiLlwACAiCjsVh75QaisAlBvuioEBfCPOH71901XpFQMAIiI
                wmzFkfsD8hiARtNVoVEOoOfi+Md+bLoipWAAQEQUVr85ci9E5HEALaarQpM4UJyDE1bfY7oixeJmQERE
                YXT/R3dDRB4GG/+wsiD4IVbOP9N0RYq/ASIiCpc/fKwJNZGHALSbrgq5ikD17moNAmpMV4CIiCYYGHof
                LCw1XQ3ySZ06/O6YaTj2wQHTVSEiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi
                IiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi
                IiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi
                IiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi
                IiIiIiIiIiIiIiIiIqpy/x9hubBNhHfHDQAAAABJRU5ErkJggg==""";

  @Autowired
  public BirtReportingProcessServiceImpl(
      final PlatformSecurityContext context,
      final IReportEngine reportEngine,
      final @Qualifier("hikariTenantDataSource") DataSource tenantDataSource,
      DatabasePasswordEncryptor databasePasswordEncryptor,
      FineractProperties fineractProperties,
      ApplicationContext applicationContext,
      ApplicationContext contextVar) {
    this.reportEngine = reportEngine;
    this.tenantDataSource = tenantDataSource;
    this.databasePasswordEncryptor = databasePasswordEncryptor;
    this.fineractProperties = fineractProperties;
    this.context = context;
    this.applicationContext = applicationContext;
    this.contextVar = contextVar;
  }

  private void updateSubReportDataSources(ReportDesignHandle designHandle) {
    List<LibraryHandle> libraries = designHandle.getAllLibraries();
    logger.debug(
        "updateSubReportDataSources() called. Library count: {}",
        libraries != null ? libraries.size() : 0);
    if (libraries != null) {
      for (LibraryHandle library : libraries) {
        setConnectionDetail(library);
        // Recursively process nested libraries
        updateNestedLibraries(library);
      }
    }
  }

  private void updateNestedLibraries(LibraryHandle libraryHandle) {
    List<LibraryHandle> nestedLibraries = libraryHandle.getAllLibraries();
    logger.debug(
        "updateNestedLibraries() called. Nested library count: {}",
        nestedLibraries != null ? nestedLibraries.size() : 0);
    if (nestedLibraries != null) {
      for (LibraryHandle nested : nestedLibraries) {
        setConnectionDetail(nested);
        updateNestedLibraries(nested);
      }
    }
  }

  @Override
  public Response processRequest(
      final String reportName, final MultivaluedMap<String, String> queryParams) {
    final var outputTypeParam = queryParams.getFirst("output-type");
    final var reportParams = getReportParams(queryParams);
    final var locale = ApiParameterHelper.extractLocale(queryParams);
    final var language = "en";

    var outputType = "HTML";
    if (StringUtils.isNotBlank(outputTypeParam)) {
      outputType = outputTypeParam;
    }

    if ((!outputType.equalsIgnoreCase("HTML")
        && !outputType.equalsIgnoreCase("PDF")
        && !outputType.equalsIgnoreCase("XLS")
        && !outputType.equalsIgnoreCase("XLSX")
        && !outputType.equalsIgnoreCase("CSV"))) {
      throw new PlatformDataIntegrityException(
          "error.msg.invalid.outputType", "No matching Output Type: " + outputType);
    }
    logger.info(
        "Processing BIRT report: name='{}', outputType='{}', locale='{}'",
        reportName,
        outputType,
        locale);

    String reportPath;
    logger.debug("locale {}", locale);
    logger.debug("language {}", language);
    if (locale != null && !"en".equalsIgnoreCase(locale.toString())) {
      reportPath =
          getReportPath() + reportName + "_" + locale.toString().toLowerCase() + ".rptdesign";
    } else {
      reportPath = getReportPath() + reportName + ".rptdesign";
    }
    logger.debug("Report path: {}", reportPath);

    // load report definition
    IReportRunnable design;

    try {
      logger.info("Attempting to load report design from path: {}", reportPath);
      if (!new File(reportPath).exists()) {
        logger.error("Report design file not found at path: {}", reportPath);
        throw new PlatformDataIntegrityException(
            "error.msg.reporting.error", "Report file not found: " + reportPath);
      }
      design = reportEngine.openReportDesign(reportPath);
      logger.info("Report design loaded successfully: '{}'", reportPath);
      final var designHandle = (ReportDesignHandle) design.getDesignHandle();

      // Override Data Connection with tenant details
      setConnectionDetail(designHandle);
      logger.debug("Main report datasource connection details updated");

      // Inject central logo into the report design
      injectCentralLogo(designHandle);

      // Update subreport data sources
      updateSubReportDataSources(designHandle);
      logger.debug("Subreport datasource connection details updated");

      // Set Locale for the report
      final var task = reportEngine.createRunAndRenderTask(design);

      // Force fast-failure on major errors
      task.setErrorHandlingOption(IEngineTask.CANCEL_ON_ERROR);
      logger.debug("BIRT task created, error handling set to CANCEL_ON_ERROR");

      try {
        if (StringUtils.isNotBlank(fineractBirtLocale)) {
          Locale localeReport = new Locale.Builder().setLanguageTag(fineractBirtLocale).build();
          task.setLocale(localeReport);
        } else if (locale != null) {
          task.setLocale(locale);
        }

        addParametersToReport(task, reportParams);
        logger.debug("Parameters bound to task successfully for report '{}'", reportName);

        final var baos = new ByteArrayOutputStream();

        if ("PDF".equalsIgnoreCase(outputType)) {
          PDFRenderOption pdfOptions = new PDFRenderOption();
          pdfOptions.setOutputFormat("pdf");
          pdfOptions.setOption(IPDFRenderOption.PAGE_OVERFLOW, IPDFRenderOption.FIT_TO_PAGE_SIZE);
          pdfOptions.setOutputStream(baos);
          task.setRenderOption(pdfOptions);
          task.run();
          logger.debug(
              "task.run() completed for report '{}', outputType='{}'", reportName, outputType);
          verifyTaskSuccess(task, reportName);
          logger.info(
              "Report '{}' generated successfully. Output size: {} bytes, type: '{}'",
              reportName,
              baos.size(),
              outputType);
          return Response.ok().entity(baos.toByteArray()).type("application/pdf").build();

        } else if ("XLS".equalsIgnoreCase(outputType)) {
          EXCELRenderOption excelOptions = new EXCELRenderOption();
          excelOptions.setOutputFormat("xls_spudsoft");
          excelOptions.setOutputStream(baos);

          // Use Spudsoft XLS emitter to support images, bypassing limitations in BIRT's native XLS
          // emitter
          excelOptions.setEmitterID("uk.co.spudsoft.birt.emitters.excel.XlsEmitter");

          // Prevents headers from disappearing when using Spudsoft emitter
          excelOptions.setOption(ExcelEmitter.STRUCTURED_HEADER, true);

          // Keeps report data in a single continuous sheet for easier filtering/sorting
          excelOptions.setOption(ExcelEmitter.SINGLE_SHEET, true);

          task.setRenderOption(excelOptions);
          task.run();
          logger.debug(
              "task.run() completed for report '{}', outputType='{}'", reportName, outputType);
          verifyTaskSuccess(task, reportName);
          logger.info(
              "Report '{}' generated successfully. Output size: {} bytes, type: '{}'",
              reportName,
              baos.size(),
              outputType);
          return Response.ok()
              .entity(baos.toByteArray())
              .type("application/vnd.ms-excel")
              .header(
                  "Content-Disposition",
                  "attachment;filename=" + reportName.replaceAll(" ", "") + ".xls")
              .build();

        } else if ("XLSX".equalsIgnoreCase(outputType)) {
          EXCELRenderOption excelOptions = new EXCELRenderOption();
          excelOptions.setOutputFormat("xlsx");
          excelOptions.setOutputStream(baos);
          task.setRenderOption(excelOptions);
          task.run();
          logger.debug(
              "task.run() completed for report '{}', outputType='{}'", reportName, outputType);
          verifyTaskSuccess(task, reportName);
          logger.info(
              "Report '{}' generated successfully. Output size: {} bytes, type: '{}'",
              reportName,
              baos.size(),
              outputType);
          return Response.ok()
              .entity(baos.toByteArray())
              .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
              .header(
                  "Content-Disposition",
                  "attachment;filename=" + reportName.replaceAll(" ", "") + ".xlsx")
              .build();

        } else if ("CSV".equalsIgnoreCase(outputType)) {
          RenderOption csvOptions = new RenderOption();
          csvOptions.setOutputFormat("csv");
          csvOptions.setOutputStream(baos);
          task.setRenderOption(csvOptions);
          task.run();
          logger.debug(
              "task.run() completed for report '{}', outputType='{}'", reportName, outputType);
          verifyTaskSuccess(task, reportName);
          logger.info(
              "Report '{}' generated successfully. Output size: {} bytes, type: '{}'",
              reportName,
              baos.size(),
              outputType);
          return Response.ok()
              .entity(baos.toByteArray())
              .type("text/csv")
              .header(
                  "Content-Disposition",
                  "attachment;filename=" + reportName.replaceAll(" ", "") + ".csv")
              .build();

        } else if ("HTML".equalsIgnoreCase(outputType)) {
          HTMLRenderOption htmlOptions = new HTMLRenderOption();
          htmlOptions.setOutputFormat("html");
          htmlOptions.setEmbeddable(true);
          htmlOptions.setOutputStream(baos);

          htmlOptions.setSupportedImageFormats("PNG");

          // Custom image handler to embed images as Base64 (Data URL) to avoid broken images caused
          // by BIRT's default temp file handling
          htmlOptions.setImageHandler(
              new HTMLServerImageHandler() {
                @Override
                protected String handleImage(
                    IImage image, Object context, String prefix, boolean needMap) {
                  if (image == null) return "";

                  try {
                    // Read the image data into a byte array
                    byte[] imageBytes = null;

                    // Prioritize direct byte access for injected images
                    if (image.getImageData() != null) {
                      imageBytes = image.getImageData();
                    }

                    // Fallback: Read from stream if no direct memory buffer
                    else if (image.getImageStream() != null) {
                      imageBytes = IOUtils.toByteArray(image.getImageStream());
                    }

                    if (imageBytes == null || imageBytes.length == 0) {
                      logger.warn(
                          "Image data is empty for image with MIME type '{}'", image.getMimeType());
                      return "";
                    }

                    // Encode the image bytes to Base64
                    String imageString = Base64.getEncoder().encodeToString(imageBytes);

                    // Return a Data URL that embeds the image directly in the HTML
                    return MessageFormat.format(
                        "data:{0};base64,{1}", image.getMimeType(), imageString);
                  } catch (IOException e) {
                    logger.error("Error embedding image in HTML output", e);
                    return "";
                  }
                }
              });

          task.setRenderOption(htmlOptions);
          task.run();
          logger.debug(
              "task.run() completed for report '{}', outputType='{}'", reportName, outputType);
          verifyTaskSuccess(task, reportName);
          logger.info(
              "Report '{}' generated successfully. Output size: {} bytes, type: '{}'",
              reportName,
              baos.size(),
              outputType);
          return Response.ok().entity(baos.toByteArray()).type("text/html").build();

        } else {
          throw new PlatformDataIntegrityException(
              "error.msg.invalid.outputType", "No matching Output Type: " + outputType);
        }
      } finally {
        task.close();
      }
    } catch (Exception e) {
      logger.error("error.msg.reporting.error:", e);
      throw new PlatformDataIntegrityException("error.msg.reporting.error", e.getMessage());
    }
  }

  private void addParametersToReport(
      final IRunAndRenderTask task, final Map<String, String> queryParams) {
    final var currentUser = this.context.authenticatedUser();
    logger.debug("addParametersToReport() called for report, user='{}'", currentUser.getUsername());
    try {
      final IGetParameterDefinitionTask paramTask =
          reportEngine.createGetParameterDefinitionTask(task.getReportRunnable());
      logger.debug("Parameter definition task created");
      try {
        for (final Object paramDefObj : paramTask.getParameterDefns(false)) {
          final IParameterDefn paramDefEntry = (IParameterDefn) paramDefObj;
          final var paramName = paramDefEntry.getName();

          logger.debug("paramName: {}", paramName);

          // Skip parameters that are injected server-side after this loop
          if (!paramName.equals("tenantUrl")
              && !paramName.equals("userhierarchy")
              && !paramName.equals("username")
              && !paramName.equals("password")
              && !paramName.equals("userid")) {

            final var pValue = queryParams.get(paramName);

            if (StringUtils.isBlank(pValue)) {
              throw new PlatformDataIntegrityException(
                  "error.msg.reporting.error", "BIRT Parameter: " + paramName + " - not Provided");
            }

            final int dataType = paramDefEntry.getDataType();
            logger.debug("addParametersToReport({} : {} : {})", paramName, pValue, dataType);

            if (dataType == IParameterDefn.TYPE_INTEGER) {
              task.setParameterValue(paramName, Integer.parseInt(pValue));
            } else if (dataType == IParameterDefn.TYPE_FLOAT
                || dataType == IParameterDefn.TYPE_DECIMAL) {
              task.setParameterValue(paramName, Double.parseDouble(pValue));
            } else if (dataType == IParameterDefn.TYPE_DATE
                || dataType == IParameterDefn.TYPE_DATE_TIME) {
              logger.debug("ParamName: {}", paramName);
              logger.debug("ParamValue: {}", pValue);
              SimpleDateFormat sdf = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
              Date date = sdf.parse(pValue);
              long millis = date.getTime();
              java.sql.Date mySQLDate = new java.sql.Date(millis);
              task.setParameterValue(paramName, mySQLDate);
              // Logging the parsed date value for debugging
              logger.debug("Date parameter '{}' parsed and set to: {}", paramName, mySQLDate);
            } else if (dataType == IParameterDefn.TYPE_BOOLEAN) {
              task.setParameterValue(paramName, Boolean.parseBoolean(pValue));
            } else {
              logger.debug("ParamName Unknown: {}", paramName);
              logger.debug("ParamValue Unknown: {}", pValue);
              task.setParameterValue(paramName, pValue);
            }
          }
        }
      } finally {
        paramTask.close();
      }

      // Context parameters for multitenant reporting
      final var tenant = ThreadLocalContextUtil.getTenant();
      final var tenantConnection = tenant.getConnection();
      String protocol = toProtocol(this.tenantDataSource);
      Environment environment = contextVar.getEnvironment();
      String tenantUrl =
          toJdbcUrl(
              protocol,
              tenantConnection.getSchemaServer(),
              tenantConnection.getSchemaServerPort(),
              tenantConnection.getSchemaName(),
              tenantConnection.getSchemaConnectionParameters());
      logger.debug("Tenant JDBC URL resolved: '{}'", tenantUrl);

      final var userhierarchy = currentUser.getOffice().getHierarchy();
      logger.debug("userhierarchy {}", userhierarchy);

      task.setParameterValue("userhierarchy", userhierarchy);

      final var userid = currentUser.getId();
      task.setParameterValue("userid", userid);

      task.setParameterValue("tenantUrl", tenantUrl.trim());

      String username;
      if (tenantConnection.getSchemaUsername() == null
          || tenantConnection.getSchemaUsername().isEmpty()) {
        username = environment.getProperty("FINERACT_DEFAULT_TENANTDB_UID");
      } else {
        username = tenantConnection.getSchemaUsername().trim();
      }
      task.setParameterValue("username", username);

      String password;
      if (tenantConnection.getSchemaPassword() == null
          || tenantConnection.getSchemaPassword().isEmpty()) {
        password = environment.getProperty("FINERACT_DEFAULT_TENANTDB_PWD");
      } else {
        password = databasePasswordEncryptor.decrypt(tenantConnection.getSchemaPassword()).trim();
      }
      task.setParameterValue("password", password);

    } catch (Exception e) {
      logger.error("error.msg.reporting.error:", e);
      throw new PlatformDataIntegrityException("error.msg.reporting.error", e.getMessage());
    }
  }

  @Override
  public Map<String, String> getReportParams(final MultivaluedMap<String, String> queryParams) {
    final Map<String, String> reportParams = new HashMap<>();
    final var keys = queryParams.keySet();
    String pKey;
    String pValue;
    for (final String k : keys) {
      if (k.startsWith("R_")) {
        pKey = k.substring(2);
        pValue = queryParams.get(k).get(0);
        reportParams.put(pKey, pValue);
      }
    }
    return reportParams;
  }

  private String getReportPath() {
    if (StringUtils.isNotBlank(fineractBirtBaseDir)) {
      return this.fineractBirtBaseDir.endsWith(File.separator)
          ? this.fineractBirtBaseDir
          : this.fineractBirtBaseDir + File.separator;
    }
    return this.mifosBaseDir + File.separator + "birtReports" + File.separator;
  }

  // Base64-encoded PNG image data for the central logo
  private void injectCentralLogo(ReportDesignHandle designHandle) {
    try {
      // remove any whitespace characters from the base64 string, which can cause decoding issues
      String cleanBase64 = CENTRAL_LOGO_BASE64.replaceAll("\\s+", "");

      // Decode the cleaned base64 string to get the image bytes
      byte[] imageBytes = java.util.Base64.getDecoder().decode(cleanBase64);

      EmbeddedImage newImage = StructureFactory.createEmbeddedImage();
      newImage.setName("mifos_logo_icon_170951.png");
      newImage.setType(DesignChoiceConstants.IMAGE_TYPE_IMAGE_PNG);
      newImage.setData(imageBytes);

      List<EmbeddedImage> imagesToRemove = new ArrayList<>();
      Iterator<?> iterator = designHandle.imagesIterator();

      // Iterate through existing images to find any with the same name and mark them for removal
      while (iterator.hasNext()) {
        EmbeddedImageHandle imgHandle = (EmbeddedImageHandle) iterator.next();
        if ("mifos_logo_icon_170951.png".equals(imgHandle.getName())) {
          imagesToRemove.add((EmbeddedImage) imgHandle.getStructure());
        }
      }

      if (!imagesToRemove.isEmpty()) {
        designHandle.dropImage(imagesToRemove);
      }

      designHandle.addImage(newImage);
      logger.info("Central logo injected successfully into report design");

    } catch (SemanticException e) {
      logger.warn("Could not inject central logo into report design: {}", e.getMessage());
      logger.warn("SemanticException details:", e);
    }
  }

  private void setConnectionDetail(ReportDesignHandle designHandle) {
    setConnectionDetailOnDataSources(designHandle.getDataSources());
  }

  private void setConnectionDetail(LibraryHandle libraryHandle) {
    setConnectionDetailOnDataSources(libraryHandle.getDataSources());
  }

  private void setConnectionDetailOnDataSources(SlotHandle dataSources) {
    logger.debug("setConnectionDetailOnDataSources() called");
    final FineractPlatformTenant tenant = ThreadLocalContextUtil.getTenant();
    final FineractPlatformTenantConnection tenantConnection = tenant.getConnection();

    Iterator<DesignElementHandle> iterator = dataSources.iterator();

    String url = getTenantUrl();
    Environment environment = contextVar.getEnvironment();

    String user;
    if (tenantConnection.getSchemaUsername() == null
        || tenantConnection.getSchemaUsername().isEmpty()) {
      user = environment.getProperty("FINERACT_DEFAULT_TENANTDB_UID");
    } else {
      user = tenantConnection.getSchemaUsername().trim();
    }

    String password;
    if (tenantConnection.getSchemaPassword() == null
        || tenantConnection.getSchemaPassword().isEmpty()) {
      password = environment.getProperty("FINERACT_DEFAULT_TENANTDB_PWD");
    } else {
      password = databasePasswordEncryptor.decrypt(tenantConnection.getSchemaPassword()).trim();
    }

    while (iterator.hasNext()) {
      Object obj = iterator.next();
      if (obj instanceof OdaDataSourceHandle dataSource) {
        try {
          dataSource.setProperty("odaURL", url);
          dataSource.setProperty("odaUser", user);
          dataSource.setProperty("odaPassword", password);
          logger.debug("Updated DataSource: {}", dataSource.getName());
        } catch (Exception e) {
          logger.error("Failed to update DataSource: " + dataSource.getName(), e);
        }
      }
    }
  }

  private String getTenantUrl() {
    final FineractPlatformTenant tenant = ThreadLocalContextUtil.getTenant();
    final FineractPlatformTenantConnection tenantConnection = tenant.getConnection();
    String protocol = toProtocol(tenantDataSource);
    // Default properties for Writing
    String schemaServer = tenantConnection.getSchemaServer();
    String schemaPort = tenantConnection.getSchemaServerPort();
    String schemaName = tenantConnection.getSchemaName();
    String schemaConnectionParameters = tenantConnection.getSchemaConnectionParameters();
    // Properties to ReadOnly case
    if (fineractProperties.getMode().isReadOnlyMode()) {
      schemaServer =
          getPropertyValue(
              tenantConnection.getReadOnlySchemaServer(),
              TenantConstants.PROPERTY_RO_SCHEMA_SERVER_NAME,
              schemaServer);
      schemaPort =
          getPropertyValue(
              tenantConnection.getReadOnlySchemaServerPort(),
              TenantConstants.PROPERTY_RO_SCHEMA_SERVER_PORT,
              schemaPort);
      schemaName =
          getPropertyValue(
              tenantConnection.getReadOnlySchemaName(),
              TenantConstants.PROPERTY_RO_SCHEMA_SCHEMA_NAME,
              schemaName);
      schemaConnectionParameters =
          getPropertyValue(
              tenantConnection.getReadOnlySchemaConnectionParameters(),
              TenantConstants.PROPERTY_RO_SCHEMA_CONNECTION_PARAMETERS,
              schemaConnectionParameters);
    }
    String jdbcUrl =
        toJdbcUrl(protocol, schemaServer, schemaPort, schemaName, schemaConnectionParameters);
    logger.debug("{}", jdbcUrl);

    return jdbcUrl;
  }

  private String getPropertyValue(
      final String baseValue, final String propertyName, final String defaultValue) {
    if (null != baseValue) {
      return baseValue;
    }
    if (applicationContext == null) {
      return defaultValue;
    }
    return applicationContext.getEnvironment().getProperty(propertyName, defaultValue);
  }

  @Override
  public List<ReportExportType> getAvailableExportTargets() {
    throw new UnsupportedOperationException("Not supported yet.");
  }

  private void verifyTaskSuccess(final IRunAndRenderTask task, final String reportName) {
    logger.debug("verifyTaskSuccess() called for report '{}'", reportName);
    final List<?> taskErrors = task.getErrors();

    if (taskErrors != null && !taskErrors.isEmpty()) {
      // Log all errors for debugging
      for (Object error : taskErrors) {
        if (error instanceof Throwable throwable) {
          logger.error(
              "BIRT internal error during report '{}': {}",
              reportName,
              throwable.getMessage(),
              throwable);
        } else {
          logger.error("BIRT internal error during report '{}': {}", reportName, error);
        }
      }

      String firstErrorMsg =
          taskErrors.get(0) instanceof Throwable
              ? ((Throwable) taskErrors.get(0)).getMessage()
              : taskErrors.get(0).toString();

      throw new PlatformDataIntegrityException(
          "error.msg.reporting.error",
          "Report generation completed with internal errors for: "
              + reportName
              + ". Starting of error: "
              + firstErrorMsg);
    }

    final int taskStatus = task.getStatus();
    if (taskStatus != IEngineTask.STATUS_SUCCEEDED) {
      logger.error("BIRT task did not succeed for report '{}'. Status: {}", reportName, taskStatus);
      throw new PlatformDataIntegrityException(
          "error.msg.reporting.error", "Report generation failed. Task status: " + taskStatus);
    }
  }
}
